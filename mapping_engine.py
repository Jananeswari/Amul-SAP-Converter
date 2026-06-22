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
    """Appends a brand-new product mapping row to the master file.
    Use this only when the SAP Code does not exist yet — for adding a
    platform SKU to an SAP Code that already exists, use
    update_or_add_mapping instead."""
    df_wide = load_master_wide(file_path)
    df_wide = pd.concat([df_wide, pd.DataFrame([new_row])], ignore_index=True)
    save_master_mapping(df_wide, file_path)
    return df_wide


def update_or_add_mapping(sap_code: str, sap_desc: str, fg_group: str,
                           platform_inputs: dict, file_path: str = None) -> dict:
    """
    Adds a new SAP product, OR — if the SAP Code already exists — attaches
    new platform SKU codes to that existing row.

    A platform SKU is "new" for that SAP code if the platform's column on
    that row is currently empty. If the platform column already holds a
    DIFFERENT SKU, that platform is reported as a conflict and is NOT
    overwritten, so existing mappings can never be silently clobbered.

    Returns a dict:
        {
          "action": "added_new_product" | "updated_existing" | "no_change",
          "updated_platforms": [...],   # platforms newly filled in
          "conflicts": {platform: existing_sku, ...},  # platforms that already had a different SKU
        }
    """
    df_wide = load_master_wide(file_path)
    platform_inputs = {p: v.strip() for p, v in platform_inputs.items() if v and v.strip()}

    existing_mask = df_wide["SAP Code"].astype(str).str.strip().str.upper() == str(sap_code).strip().upper()

    if not existing_mask.any():
        new_row = {
            "FG Group": "", "FG Group Description": fg_group,
            "Article Description ": sap_desc, "GTIN": None,
            "SAP Code": sap_code, "Product Description as per SAP": sap_desc,
        }
        new_row.update(platform_inputs)
        df_wide = pd.concat([df_wide, pd.DataFrame([new_row])], ignore_index=True)
        save_master_mapping(df_wide, file_path)
        return {"action": "added_new_product", "updated_platforms": list(platform_inputs.keys()), "conflicts": {}}

    row_idx = df_wide[existing_mask].index[0]
    updated_platforms, conflicts = [], {}

    for platform, new_sku in platform_inputs.items():
        if platform not in df_wide.columns:
            df_wide[platform] = pd.NA
        current_val = df_wide.at[row_idx, platform]
        current_val_str = str(current_val).strip() if pd.notna(current_val) else ""

        if not current_val_str:
            df_wide.at[row_idx, platform] = new_sku
            updated_platforms.append(platform)
        elif current_val_str.upper() != new_sku.upper():
            conflicts[platform] = current_val_str
        # if it's already the same SKU, no change needed

    if updated_platforms:
        save_master_mapping(df_wide, file_path)
        return {"action": "updated_existing", "updated_platforms": updated_platforms, "conflicts": conflicts}

    return {"action": "no_change", "updated_platforms": [], "conflicts": conflicts}


def bulk_update_from_unmapped_list(filled_df: pd.DataFrame, file_path: str = None) -> dict:
    """
    Takes the manager's filled-in "Unmapped SKUs" sheet — same shape as the
    download (Platform | SKU | Product Name (from platform) | Qty Ordered),
    plus a SAP Code column the manager has typed in — and applies the same
    safe rule as the single-entry form to every row:

      - if the SAP Code doesn't exist at all -> that row is skipped and
        reported (bulk update only attaches SKUs to EXISTING SAP codes,
        since there's no description/FG group supplied here)
      - if the platform column on that SAP Code's row is empty -> filled in
      - if it already holds a DIFFERENT SKU -> left untouched, reported as
        a conflict
      - if it already holds the SAME SKU -> skipped silently

    Expected columns in filled_df (case-insensitive, whitespace-tolerant):
        Platform, SKU, SAP Code
    (Product Name / Qty Ordered columns, if present, are ignored.)

    Returns a summary dict:
        {
          "updated_count": int,
          "skipped_no_sap_code": [...],      # rows where SAP Code cell was left blank
          "skipped_unknown_sap_code": [...], # rows whose typed SAP Code doesn't exist in master
          "conflicts": [...],                # rows where the platform already had a different SKU
          "updated_rows": [...],             # rows successfully applied
        }
    """
    cols_lower = {c.strip().lower(): c for c in filled_df.columns}
    required = ["platform", "sku", "sap code"]
    missing = [r for r in required if r not in cols_lower]
    if missing:
        raise ValueError(
            f"Uploaded file is missing required column(s): {', '.join(missing)}. "
            f"Found columns: {list(filled_df.columns)}"
        )

    platform_col = cols_lower["platform"]
    sku_col = cols_lower["sku"]
    sap_col = cols_lower["sap code"]

    df_wide = load_master_wide(file_path)
    sap_lookup = {
        str(code).strip().upper(): idx
        for idx, code in df_wide["SAP Code"].items() if pd.notna(code)
    }

    updated_count = 0
    skipped_no_sap_code, skipped_unknown_sap_code, conflicts, updated_rows = [], [], [], []

    for _, row in filled_df.iterrows():
        platform = str(row[platform_col]).strip() if pd.notna(row[platform_col]) else ""
        sku = str(row[sku_col]).strip() if pd.notna(row[sku_col]) else ""
        sap_code_in = str(row[sap_col]).strip() if pd.notna(row[sap_col]) else ""

        if not sku or not platform:
            continue  # nothing to do for a blank platform/sku row

        if not sap_code_in:
            skipped_no_sap_code.append({"platform": platform, "sku": sku})
            continue

        row_idx = sap_lookup.get(sap_code_in.upper())
        if row_idx is None:
            skipped_unknown_sap_code.append({"platform": platform, "sku": sku, "sap_code_entered": sap_code_in})
            continue

        if platform not in df_wide.columns:
            df_wide[platform] = pd.NA
        current_val = df_wide.at[row_idx, platform]
        current_val_str = str(current_val).strip() if pd.notna(current_val) else ""

        if not current_val_str:
            df_wide.at[row_idx, platform] = sku
            updated_count += 1
            updated_rows.append({"platform": platform, "sku": sku, "sap_code": sap_code_in})
        elif current_val_str.upper() != sku.upper():
            conflicts.append({"platform": platform, "sku": sku, "sap_code": sap_code_in, "existing_sku": current_val_str})
        # else: identical SKU already mapped -> nothing to do

    if updated_count:
        save_master_mapping(df_wide, file_path)

    return {
        "updated_count": updated_count,
        "skipped_no_sap_code": skipped_no_sap_code,
        "skipped_unknown_sap_code": skipped_unknown_sap_code,
        "conflicts": conflicts,
        "updated_rows": updated_rows,
    }
