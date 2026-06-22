"""
Core Mapping Engine
Loads the master Article Master file, parses case-pack sizes from SAP
descriptions, and provides lookup utilities used by the ETL pipeline.
"""

import re
import pandas as pd
import os

MASTER_FILE_PATH = os.path.join(os.path.dirname(__file__), "data", "Amul_Article_Master.xlsx")

# Platform columns present in the master file (order matters for display)
PLATFORM_COLUMNS = [
    "Amazon", "Bigbasket", "D mart", "Flipkart", "Metro",
    "Reliance", "Swiggy", "Zepto", "Star", "Lots", "Spar", "Blinkit",
]


def parse_case_pack(description: str):
    """
    Extracts the number of consumer units packed into one SAP case/box
    from a free-text SAP product description.
    Returns (case_pack_size: int, confidence: str)
    confidence is one of: "high", "low" (ambiguous combo pack), "default" (assumed 1).
    """
    if not isinstance(description, str) or not description.strip():
        return 1, "default"
    d = description.strip()

    # Ambiguous combo packs like (1+1f)X19) -- flag as low confidence, don't guess silently
    if re.search(r'\(\s*\d+\s*\+\s*\d+\s*f?\s*\)', d, re.IGNORECASE):
        m = re.search(r'\(\s*\d+\s*\+\s*\d+\s*f?\s*\)\s*[xX]\s*(\d+)', d, re.IGNORECASE)
        if m:
            return int(m.group(1)), "low"
        return 1, "low"

    # N x ( N x ... )  e.g. 8x(6x200g)
    m = re.search(r'(\d+)\s*[xX]\s*\(\s*(\d+)\s*[xX*]', d)
    if m:
        return int(m.group(1)) * int(m.group(2)), "high"

    # N x N x num unit   e.g. 8x8x200 Gm, 9x20x50 Gm
    m = re.search(r'(\d+)\s*[xX]\s*(\d+)\s*[xX]\s*[\d.]+\s*[a-zA-Z]', d)
    if m:
        return int(m.group(1)) * int(m.group(2)), "high"

    # N x num(.num) optional unit   e.g. 30x200ml, 24 x 500 ml, 12x1 Litre, 60x200
    m = re.search(r'(\d+)\s*[xX*]\s*[\d.]+\s*[a-zA-Z]*', d)
    if m:
        return int(m.group(1)), "high"

    # trailing (1x10) or (3x6) style — case x inner
    m = re.search(r'\((\d+)\s*[xX]\s*(\d+)\)', d)
    if m:
        return int(m.group(1)) * int(m.group(2)), "high"

    # No pack pattern found -> assume sold as single unit (case pack = 1)
    return 1, "default"


def load_master_mapping(file_path: str = None) -> pd.DataFrame:
    """
    Loads and normalizes the Article Master file.
    Returns a long-format DataFrame: one row per (SAP Code, Platform, Platform SKU).
    """
    path = file_path or MASTER_FILE_PATH
    df = pd.read_excel(path, sheet_name="Sheet1")
    df.columns = [c.strip() for c in df.columns]

    df = df.dropna(subset=["SAP Code"]).copy()
    df["case_pack_size"], df["pack_confidence"] = zip(
        *df["Product Description as per SAP"].apply(parse_case_pack)
    )

    long_rows = []
    for _, row in df.iterrows():
        for platform in PLATFORM_COLUMNS:
            sku_val = row.get(platform)
            if pd.notna(sku_val) and str(sku_val).strip():
                long_rows.append({
                    "platform": platform,
                    "platform_sku": str(sku_val).strip().upper(),
                    "sap_code": row["SAP Code"],
                    "sap_description": row["Product Description as per SAP"],
                    "fg_group": row.get("FG Group Description", ""),
                    "case_pack_size": row["case_pack_size"],
                    "pack_confidence": row["pack_confidence"],
                })
    return pd.DataFrame(long_rows)


def save_master_mapping(df_wide: pd.DataFrame, file_path: str = None):
    """Saves the wide-format master mapping back to disk (used by the admin 'add mapping' form)."""
    path = file_path or MASTER_FILE_PATH
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df_wide.to_excel(writer, sheet_name="Sheet1", index=False)


def load_master_wide(file_path: str = None) -> pd.DataFrame:
    """Loads the original wide-format master file (one row per product, platform columns side by side)."""
    path = file_path or MASTER_FILE_PATH
    df = pd.read_excel(path, sheet_name="Sheet1")
    df.columns = [c.strip() for c in df.columns]
    return df


def add_new_mapping_row(new_row: dict, file_path: str = None):
    """Appends a new product mapping row to the master file."""
    df_wide = load_master_wide(file_path)
    df_wide = pd.concat([df_wide, pd.DataFrame([new_row])], ignore_index=True)
    save_master_mapping(df_wide, file_path)
    return df_wide
