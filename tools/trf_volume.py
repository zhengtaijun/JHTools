import re
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import streamlit as st

from utils.excel_io import read_excel_any
from utils.product_matcher import match_cbm


def render():
    st.subheader("📦 TRF Volume Calculator")
    st.markdown("📺 [Instructional video](https://youtu.be/S10a3kPEXZg)")

    warehouse_file = st.file_uploader("Upload warehouse export (Excel)", type=["xlsx", "xls"])

    df_preview = None
    uploaded_bytes = None
    n_cols = 0

    if warehouse_file:
        try:
            uploaded_bytes = warehouse_file.getvalue()
            df_preview = read_excel_any(BytesIO(uploaded_bytes))
            n_cols = df_preview.shape[1]

            with st.expander("📋 Preview uploaded file & columns", expanded=False):
                st.write(df_preview.head())
                st.write("Columns:", list(df_preview.columns))
        except Exception as e:
            st.error(f"❌ Failed to read uploaded file: {e}")

    st.markdown("### 🔠 Column headers (auto-detect, editable)")
    header_prod = st.text_input("Header for **Product Name**", value="Short Description")
    header_order = st.text_input("Header for **Order Number** (Invoices)", value="Invoices")
    header_qty = st.text_input("Header for **Quantity**", value="Order Qty")
    header_po = st.text_input("Header for **PO Number**", value="PO No")

    def _auto_detect_col(df, header_name, fallback=1):
        if df is None or not header_name:
            return fallback
        cols_lower = [str(c).strip().lower() for c in df.columns]
        target = header_name.strip().lower()
        for i, c in enumerate(cols_lower):
            if c == target:
                return i + 1
        for i, c in enumerate(cols_lower):
            if target in c:
                return i + 1
        return fallback

    max_cols = n_cols if n_cols else 50
    auto_prod = _auto_detect_col(df_preview, header_prod, fallback=3)
    auto_order = _auto_detect_col(df_preview, header_order, fallback=7)
    auto_qty = _auto_detect_col(df_preview, header_qty, fallback=8)
    auto_po = _auto_detect_col(df_preview, header_po, fallback=1)

    st.markdown("### #️⃣ Column numbers (1-based, optional override)")
    use_manual_cols = st.checkbox(
        "Use manual column numbers to override header detection",
        value=False,
        help="默认关闭：仅使用上面的表头标题来识别列。勾选后：强制使用下面的列号。"
    )

    col_prod = st.number_input("Column # of **Product Name**", 1, max_cols, value=min(max(auto_prod, 1), max_cols))
    col_order = st.number_input("Column # of **Order Number (Invoices)**", 1, max_cols, value=min(max(auto_order, 1), max_cols))
    col_qty = st.number_input("Column # of **Quantity**", 1, max_cols, value=min(max(auto_qty, 1), max_cols))
    col_po = st.number_input("Column # of **PO Number**", 1, max_cols, value=min(max(auto_po, 1), max_cols))

    def resolve_col_index(df, header_name, manual_1based, auto_1based, use_manual: bool):
        if use_manual and manual_1based is not None:
            return int(manual_1based) - 1

        if df is not None and header_name:
            cols_lower = [str(c).strip().lower() for c in df.columns]
            target = header_name.strip().lower()
            for i, c in enumerate(cols_lower):
                if c == target:
                    return i
            for i, c in enumerate(cols_lower):
                if target and target in c:
                    return i

        if auto_1based is not None:
            return int(auto_1based) - 1

        raise ValueError(f"Cannot resolve column for header '{header_name}'")

    def process_volume_file(file_bytes, prod_idx, qty_idx, inv_idx, po_idx):
        dfw = read_excel_any(BytesIO(file_bytes))

        ncols = dfw.shape[1]
        for idx_check, label in [(prod_idx, "Product Name"), (qty_idx, "Quantity"), (inv_idx, "Invoices"), (po_idx, "PO No")]:
            if idx_check < 0 or idx_check >= ncols:
                raise ValueError(f"Column index for {label} out of range.")

        product_series = dfw.iloc[:, prod_idx].fillna("").astype(str)
        product_names = product_series.tolist()
        quantities = pd.to_numeric(dfw.iloc[:, qty_idx], errors="coerce").fillna(0)

        inv_series = dfw.iloc[:, inv_idx] if inv_idx is not None else pd.Series([""] * len(dfw))
        po_series = dfw.iloc[:, po_idx] if po_idx is not None else pd.Series([""] * len(dfw))

        merged_ref = []
        for po, inv in zip(po_series, inv_series):
            po_s = "" if pd.isna(po) else str(po).strip()
            if po_s:
                if re.fullmatch(r"\d+(\.0+)?", po_s):
                    po_s = po_s.split(".", 1)[0]
                if not po_s.upper().startswith("PO"):
                    po_s = f"PO{po_s}"

            inv_s = "" if pd.isna(inv) else str(inv).strip()

            if po_s and inv_s:
                merged_ref.append(f"{po_s}, {inv_s}")
            elif po_s:
                merged_ref.append(po_s)
            else:
                merged_ref.append(inv_s)

        total = len(product_names)
        volumes = [0.0] * total

        def worker(start: int, end: int):
            out = []
            for i in range(start, end):
                nm = product_names[i].strip()
                out.append(match_cbm(nm) if nm else 0.0)
            return out

        with ThreadPoolExecutor(max_workers=4) as pool:
            chunk = max(total // 4, 1)
            futures = [pool.submit(worker, i * chunk, (i + 1) * chunk if i < 3 else total) for i in range(4)]
            pos = 0
            for f in futures:
                batch = f.result()
                volumes[pos: pos + len(batch)] = batch
                pos += len(batch)

        df_res = pd.DataFrame({
            "PO/Invoice": merged_ref,
            "Short Description": product_series,
            "Order Qty": quantities,
        })

        df_res["Volume"] = pd.to_numeric(pd.Series(volumes), errors="coerce").fillna(0)
        df_res["Total Volume"] = df_res["Volume"] * df_res["Order Qty"]

        summary = pd.DataFrame({
            "PO/Invoice": [""],
            "Short Description": [""],
            "Order Qty": [""],
            "Volume": [""],
            "Total Volume": [df_res["Total Volume"].sum()],
        })

        return pd.concat([df_res, summary], ignore_index=True)

    if warehouse_file and uploaded_bytes and st.button("Calculate volume"):
        if df_preview is None:
            st.error("❌ File not loaded correctly, please re-upload.")
            return

        with st.spinner("Processing…"):
            try:
                prod_idx = resolve_col_index(df_preview, header_prod, col_prod, auto_prod, use_manual_cols)
                qty_idx = resolve_col_index(df_preview, header_qty, col_qty, auto_qty, use_manual_cols)
                inv_idx = resolve_col_index(df_preview, header_order, col_order, auto_order, use_manual_cols)
                po_idx = resolve_col_index(df_preview, header_po, col_po, auto_po, use_manual_cols)

                result_df = process_volume_file(uploaded_bytes, prod_idx, qty_idx, inv_idx, po_idx)

                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    result_df.to_excel(writer, index=False)
                buffer.seek(0)

                st.success("✅ Volume calculation complete.")
                st.dataframe(result_df.head(50), use_container_width=True)

                st.download_button(
                    "📥 Download Excel",
                    buffer,
                    file_name="TRF_Volume_Result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"❌ Error: {e}")
