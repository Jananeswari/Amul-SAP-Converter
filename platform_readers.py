"""
Platform File Readers
Each e-commerce / quick-commerce platform exports POs in its own column
layout. These functions normalize each into a common internal schema:

    platform_sku | order_qty_units | order_date | raw_product_name

`order_date` here is the date column the business wants to plan
delivery against (recommended_date for Zepto-style files).
"""

import pandas as pd
import io


def _read_all_sheets(file_bytes) -> pd.DataFrame:
    """Reads every sheet in the uploaded workbook and concatenates them
    (Zepto-style exports often split rows across Sheet1/Sheet2 by warehouse)."""
    sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
    return pd.concat(sheets.values(), ignore_index=True)


def _excel_serial_to_date(series: pd.Series) -> pd.Series:
    """Converts Excel serial date numbers to real dates; passes through
    already-parsed dates untouched."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_datetime(series, unit="D", origin="1899-12-30", errors="coerce")
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def read_zepto_po(file_bytes) -> pd.DataFrame:
    """
    Zepto Open PO export format.
    Key columns: sku (GUID), product_name, pack_size, uom, po_qty,
    recommended_date, wh_name, city_name, postatus.
    """
    df = _read_all_sheets(file_bytes)
    df.columns = [c.strip() for c in df.columns]

    out = pd.DataFrame({
        "platform_sku":      df["sku"].astype(str).str.strip().str.upper(),
        "raw_product_name":  df.get("product_name", ""),
        "order_qty_units":   pd.to_numeric(df.get("po_qty", df.get("open_qty")), errors="coerce"),
        "order_date":        _excel_serial_to_date(df["recommended_date"]),
        "warehouse":         df.get("wh_name", ""),
        "city":              df.get("city_name", ""),
        "po_status":         df.get("postatus", ""),
        "po_reference":      df.get("externpocode", ""),
    })
    out["order_qty_units"] = out["order_qty_units"].fillna(0)
    return out.dropna(subset=["platform_sku"])


def read_generic_po(file_bytes, sku_col, qty_col, date_col, name_col=None,
                     po_ref_col=None) -> pd.DataFrame:
    """
    Generic reader for platforms whose column names we configure manually
    (used for Blinkit / Swiggy Instamart / BigBasket / Reliance Mart once
    their real export formats are shared).
    """
    df = _read_all_sheets(file_bytes)
    df.columns = [c.strip() for c in df.columns]

    out = pd.DataFrame({
        "platform_sku":      df[sku_col].astype(str).str.strip().str.upper(),
        "raw_product_name":  df[name_col] if name_col and name_col in df.columns else "",
        "order_qty_units":   pd.to_numeric(df[qty_col], errors="coerce").fillna(0),
        "order_date":        _excel_serial_to_date(df[date_col]) if date_col in df.columns else pd.NaT,
        "warehouse":         "",
        "city":              "",
        "po_status":         "",
        "po_reference":      df[po_ref_col] if po_ref_col and po_ref_col in df.columns else "",
    })
    return out.dropna(subset=["platform_sku"])


# Registry: maps platform name -> reader function + required config
PLATFORM_READERS = {
    "Zepto": {"reader": read_zepto_po, "configurable": False},
    "Blinkit": {
        "reader": read_generic_po, "configurable": True,
        "default_cols": {"sku_col": "SKU_Code", "qty_col": "Quantity_Ordered",
                          "date_col": "Order_Date", "name_col": "Product_Name",
                          "po_ref_col": "Blinkit_Order_ID"},
    },
    "Swiggy Instamart": {
        "reader": read_generic_po, "configurable": True,
        "default_cols": {"sku_col": "item_sku", "qty_col": "qty",
                          "date_col": "order_date", "name_col": "item_name",
                          "po_ref_col": "order_number"},
    },
    "BigBasket": {
        "reader": read_generic_po, "configurable": True,
        "default_cols": {"sku_col": "Vendor SKU", "qty_col": "Units Ordered",
                          "date_col": "Purchase Date", "name_col": "Product Description",
                          "po_ref_col": "BB Order Ref"},
    },
    "Reliance Mart": {
        "reader": read_generic_po, "configurable": True,
        "default_cols": {"sku_col": "Material Code", "qty_col": "Order Qty",
                          "date_col": "PO Date", "name_col": "Material Desc",
                          "po_ref_col": "PO Number"},
    },
}
