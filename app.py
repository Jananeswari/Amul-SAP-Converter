"""
Amul SAP PO Converter — Streamlit Web App
Upload e-commerce platform PO files -> Get SAP-mapped monthly box
quantity projections, ready for raising Purchase Orders.
"""

import streamlit as st
import pandas as pd
from mapping_engine import load_master_wide, update_or_add_mapping, bulk_update_from_unmapped_list, PLATFORM_COLUMNS, parse_case_pack
from platform_readers import PLATFORM_READERS, read_generic_po
from etl_engine import convert_platform_orders, build_manager_projection, export_projection_to_excel

st.set_page_config(page_title="Amul SAP PO Converter", page_icon="📦", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #FAFAFA; }
    div[data-testid="metric-container"] {
        background: white; border: 1px solid #E5E5E5; border-radius: 8px;
        padding: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .stDownloadButton > button {
        background-color: #14223B !important; color: white !important;
        border-radius: 6px !important; font-weight: 600; width: 100%;
    }
    .flag-pill {
        background: #FFF3CD; color: #946C00; padding: 2px 9px;
        border-radius: 12px; font-size: 12.5px; font-weight: 600;
    }
    .ok-pill {
        background: #E6F4EA; color: #1E7E34; padding: 2px 9px;
        border-radius: 12px; font-size: 12.5px; font-weight: 600;
    }
    h1 { color: #14223B; }
</style>
""", unsafe_allow_html=True)

if "results_cache" not in st.session_state:
    st.session_state.results_cache = {}

# ── Sidebar Navigation ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📦 Amul SAP Converter")
    page = st.radio("Go to", ["Convert Orders", "Manage SKU Mapping"], label_visibility="collapsed")
    st.markdown("---")
    st.caption(f"Master mapping file: **{len(load_master_wide())} products** loaded")

# ══════════════════════════════════════════════════════════════════════════
# PAGE 1: CONVERT ORDERS
# ══════════════════════════════════════════════════════════════════════════
if page == "Convert Orders":
    st.markdown("# Convert Platform Orders to SAP PO Quantities")
    st.caption("Upload PO files from each platform. Get back Amul SAP codes, box quantities, and order dates — ready to raise POs.")
    st.markdown("---")

    available_platforms = list(PLATFORM_READERS.keys())
    uploaded = {}

    cols = st.columns(len(available_platforms))
    for i, platform in enumerate(available_platforms):
        with cols[i]:
            st.markdown(f"**{platform}**")
            f = st.file_uploader(f"Upload {platform}", type=["xlsx", "xls"],
                                  key=f"up_{platform}", label_visibility="collapsed")
            if f:
                uploaded[platform] = f
                st.markdown('<span class="ok-pill">✓ Ready</span>', unsafe_allow_html=True)

    st.markdown("---")
    process = st.button("⚡ Convert to SAP PO Quantities", type="primary",
                         disabled=len(uploaded) == 0, use_container_width=True)

    if process and uploaded:
        all_mapped, unmapped_by_platform, stats_list = {}, {}, []
        progress = st.progress(0, text="Reading files...")

        for i, (platform, file) in enumerate(uploaded.items()):
            progress.progress((i + 1) / len(uploaded), text=f"Processing {platform}...")
            reader_cfg = PLATFORM_READERS[platform]
            file_bytes = file.read()

            if reader_cfg["configurable"]:
                cols_cfg = reader_cfg["default_cols"]
                orders = read_generic_po(file_bytes, cols_cfg["sku_col"], cols_cfg["qty_col"],
                                          cols_cfg["date_col"], cols_cfg.get("name_col"),
                                          cols_cfg.get("po_ref_col"))
            else:
                orders = reader_cfg["reader"](file_bytes)

            result = convert_platform_orders(orders, platform)
            all_mapped[platform] = result["mapped"]
            unmapped_by_platform[platform] = result["unmapped"]
            stats_list.append(result["stats"])

        progress.empty()
        st.session_state.results_cache = {
            "all_mapped": all_mapped, "unmapped": unmapped_by_platform, "stats": stats_list,
        }

    if st.session_state.results_cache:
        all_mapped = st.session_state.results_cache["all_mapped"]
        unmapped_by_platform = st.session_state.results_cache["unmapped"]
        stats_list = st.session_state.results_cache["stats"]

        st.success("✅ Conversion complete", icon="✅")
        st.markdown("## Summary")

        k1, k2, k3, k4 = st.columns(4)
        total_rows = sum(s["total_rows"] for s in stats_list)
        total_mapped = sum(s["mapped_rows"] for s in stats_list)
        total_boxes = sum(s["total_boxes"] for s in stats_list)
        total_rounded = sum(s["rounded_up_rows"] for s in stats_list)

        k1.metric("Order Lines Processed", f"{total_rows:,}")
        k2.metric("Successfully Mapped", f"{total_mapped:,}",
                  delta=f"{round(total_mapped/total_rows*100) if total_rows else 0}% match rate")
        k3.metric("Total SAP PO Boxes", f"{total_boxes:,.0f}")
        k4.metric("Rows Rounded Up", f"{total_rounded:,}", delta="check these before PO", delta_color="off")

        st.markdown("---")
        st.markdown("## Monthly PO Projection (Manager View)")
        st.caption("One row per SAP product. Quantity, date, and rounding flag shown separately for each platform.")

        projection = build_manager_projection(all_mapped)

        if not projection.empty:
            rounded_cols = [c for c in projection.columns if "Rounded Up" in c]
            highlight_mask = projection[rounded_cols].any(axis=1) if rounded_cols else pd.Series(False, index=projection.index)

            def highlight_rows(row):
                return ['background-color: #FFF3CD' if highlight_mask.loc[row.name] else '' for _ in row]

            st.dataframe(projection.style.apply(highlight_rows, axis=1), use_container_width=True, height=420)
            st.caption(f"🟨 {int(highlight_mask.sum())} rows had a remainder and were rounded up — review highlighted rows before finalizing the PO.")
        else:
            st.warning("No SKUs were successfully mapped. Check the Unmapped tab below.")

        st.markdown("---")
        st.markdown("## Per-Platform Detail")
        tabs = st.tabs(list(all_mapped.keys()) + ["⚠️ Unmapped SKUs"])
        for i, platform in enumerate(all_mapped.keys()):
            with tabs[i]:
                s = next(s for s in stats_list if s["platform"] == platform)
                c1, c2, c3 = st.columns(3)
                c1.metric("Rows", s["total_rows"])
                c2.metric("Mapped", s["mapped_rows"])
                c3.metric("Unmapped", s["unmapped_rows"])
                df_show = all_mapped[platform][[
                    "platform_sku", "sap_code", "sap_description",
                    "order_qty_units", "case_pack_size", "po_qty_boxes",
                    "had_remainder", "order_date"
                ]].rename(columns={
                    "platform_sku": "Platform SKU", "sap_code": "SAP Code",
                    "sap_description": "SAP Description", "order_qty_units": "Ordered (Units)",
                    "case_pack_size": "Case Pack Size", "po_qty_boxes": "PO Qty (Boxes)",
                    "had_remainder": "Rounded Up?", "order_date": "Order Date",
                })
                st.dataframe(df_show, use_container_width=True, height=300)

        with tabs[-1]:
            unmapped_frames = []
            for platform, df in unmapped_by_platform.items():
                if not df.empty:
                    tmp = df[["platform_sku", "raw_product_name", "order_qty_units"]].copy()
                    tmp.insert(0, "Platform", platform)
                    unmapped_frames.append(tmp)
            if unmapped_frames:
                df_unm = pd.concat(unmapped_frames, ignore_index=True)
                df_unm.columns = ["Platform", "SKU", "Product Name (from platform)", "Qty Ordered"]

                # Dedupe by Platform+SKU for the fill-in template — manager only
                # needs to type the SAP Code once per unique SKU, not once per order line
                template_df = (
                    df_unm.groupby(["Platform", "SKU"], as_index=False)
                    .agg({"Product Name (from platform)": "first", "Qty Ordered": "sum"})
                )
                template_df["SAP Code"] = ""

                st.warning(f"{len(df_unm)} order lines ({len(template_df)} unique SKUs) could not be matched to a SAP code.")
                st.caption("Fill in the **SAP Code** column below, then go to **Manage SKU Mapping → Bulk Update** in the sidebar and upload it back.")
                st.dataframe(df_unm, use_container_width=True, height=300)

                import io as _io
                buf = _io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    template_df.to_excel(writer, sheet_name="Fill SAP Code", index=False)
                st.download_button(
                    "📥 Download Fill-In Template (for Bulk Mapping Update)",
                    data=buf.getvalue(),
                    file_name="Unmapped_SKUs_Fill_SAP_Code.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.success("🎉 All SKUs were mapped successfully!")

        st.markdown("---")
        st.markdown("## Download")
        excel_bytes = export_projection_to_excel(projection, unmapped_by_platform, stats_list)
        st.download_button(
            "📥 Download SAP PO Projection (Excel)",
            data=excel_bytes,
            file_name="Amul_SAP_PO_Projection.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption("Includes: Summary | Monthly PO Projection (rounded-up rows highlighted) | Unmapped SKUs")

# ══════════════════════════════════════════════════════════════════════════
# PAGE 2: MANAGE SKU MAPPING
# ══════════════════════════════════════════════════════════════════════════
elif page == "Manage SKU Mapping":
    st.markdown("# Manage SKU Mapping")
    st.caption("Add new products or new platform SKU codes to the master mapping table.")
    st.markdown("---")

    master_df = load_master_wide()

    tab1, tab2, tab3 = st.tabs(["➕ Add New Mapping", "📤 Bulk Update (Unmapped List)", "🔍 View / Search Mapping"])

    with tab1:
        st.markdown("### Add a new product, or attach a platform SKU to an existing SAP Code")
        st.caption("If the SAP Code already exists, leave description/FG group blank — just fill in the new platform SKU(s) and they'll be attached to that existing product.")

        with st.form("add_mapping_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                sap_code = st.text_input("SAP Code *", placeholder="e.g. TDMCP01")
                sap_desc = st.text_input("SAP Product Description (required only for a NEW SAP Code)", placeholder="e.g. Amul Taaza Fresh Toned Milk 12x1 Ltr TP")
            with c2:
                fg_group = st.text_input("FG Group Description (optional)", placeholder="e.g. Milk - UHT")
                if sap_desc:
                    cp, conf = parse_case_pack(sap_desc)
                    badge = "✅ detected" if conf == "high" else ("⚠️ ambiguous, please verify" if conf == "low" else "ℹ️ defaulted to 1 (no pack pattern found)")
                    st.info(f"Case pack size: **{cp} units/box** — {badge}")

            st.markdown("**Platform SKU Codes** (fill in whichever platforms apply)")
            platform_inputs = {}
            pcols = st.columns(4)
            for i, platform in enumerate(PLATFORM_COLUMNS):
                with pcols[i % 4]:
                    platform_inputs[platform] = st.text_input(platform, key=f"pf_{platform}")

            submitted = st.form_submit_button("Save Mapping", type="primary", use_container_width=True)

            if submitted:
                sap_code_clean = sap_code.strip() if sap_code else ""
                is_existing = sap_code_clean.upper() in master_df["SAP Code"].astype(str).str.strip().str.upper().values

                if not sap_code_clean:
                    st.error("SAP Code is required.")
                elif not is_existing and not sap_desc:
                    st.error(f"SAP Code '{sap_code_clean}' doesn't exist yet — SAP Product Description is required to create it.")
                elif not any(v.strip() for v in platform_inputs.values()):
                    st.error("Enter at least one platform SKU code.")
                else:
                    result = update_or_add_mapping(sap_code_clean, sap_desc, fg_group, platform_inputs)

                    if result["action"] == "added_new_product":
                        st.success(f"✅ Added new SAP Code **{sap_code_clean}** with platform SKU(s): {', '.join(result['updated_platforms'])}")
                        st.rerun()
                    elif result["action"] == "updated_existing":
                        st.success(f"✅ SAP Code **{sap_code_clean}** already existed — attached new SKU(s) for: {', '.join(result['updated_platforms'])}")
                        if result["conflicts"]:
                            for plat, existing_sku in result["conflicts"].items():
                                st.warning(f"⚠️ {plat} already had a different SKU mapped (`{existing_sku}`) — not overwritten. Edit it in the master file directly if this needs to change.")
                        st.rerun()
                    else:  # no_change
                        st.warning("No new SKUs were added — every platform you entered already had the exact same SKU mapped, or all conflicted with an existing different SKU:")
                        for plat, existing_sku in result["conflicts"].items():
                            st.warning(f"⚠️ {plat} already maps to a different SKU (`{existing_sku}`).")

    with tab2:
        st.markdown("### Bulk update from a filled-in Unmapped SKUs list")
        st.caption(
            "Go to **Convert Orders → Unmapped SKUs tab**, download the fill-in template, "
            "type the correct **SAP Code** next to each SKU, then upload it here. "
            "Only attaches SKUs to SAP Codes that already exist — empty platform slots are filled in, "
            "existing different SKUs are never overwritten."
        )

        bulk_file = st.file_uploader("Upload filled-in template", type=["xlsx", "xls"], key="bulk_upload")

        if bulk_file:
            try:
                preview_df = pd.read_excel(bulk_file)
                st.markdown(f"**Preview** ({len(preview_df)} rows)")
                st.dataframe(preview_df, use_container_width=True, height=250)

                if st.button("Apply Bulk Update", type="primary", use_container_width=True):
                    bulk_file.seek(0)
                    apply_df = pd.read_excel(bulk_file)
                    result = bulk_update_from_unmapped_list(apply_df)

                    if result["updated_count"]:
                        st.success(f"✅ {result['updated_count']} SKU(s) successfully mapped.")
                        with st.expander("See what was added"):
                            st.dataframe(pd.DataFrame(result["updated_rows"]), use_container_width=True)
                    else:
                        st.info("No new SKUs were mapped — see details below for why.")

                    if result["conflicts"]:
                        st.warning(f"⚠️ {len(result['conflicts'])} row(s) skipped — that platform already had a DIFFERENT SKU mapped to the SAP Code given. Nothing was overwritten.")
                        st.dataframe(pd.DataFrame(result["conflicts"]), use_container_width=True)

                    if result["skipped_unknown_sap_code"]:
                        st.warning(f"⚠️ {len(result['skipped_unknown_sap_code'])} row(s) skipped — the SAP Code typed in doesn't exist in the master file yet. Use **Add New Mapping** for these instead.")
                        st.dataframe(pd.DataFrame(result["skipped_unknown_sap_code"]), use_container_width=True)

                    if result["skipped_no_sap_code"]:
                        st.info(f"ℹ️ {len(result['skipped_no_sap_code'])} row(s) skipped — SAP Code column was left blank.")
                        st.dataframe(pd.DataFrame(result["skipped_no_sap_code"]), use_container_width=True)

            except ValueError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"Couldn't read this file: {e}")

    with tab3:
        st.markdown("### Search the master mapping table")
        search = st.text_input("Search by SAP Code, Description, or FG Group", placeholder="Type to filter...")
        display_df = master_df
        if search:
            mask = (
                master_df["SAP Code"].astype(str).str.contains(search, case=False, na=False)
                | master_df["Product Description as per SAP"].astype(str).str.contains(search, case=False, na=False)
                | master_df["FG Group Description"].astype(str).str.contains(search, case=False, na=False)
            )
            display_df = master_df[mask]
        st.caption(f"Showing {len(display_df)} of {len(master_df)} products")
        st.dataframe(display_df, use_container_width=True, height=500)
