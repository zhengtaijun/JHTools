# tools/excel_column_remover.py
import streamlit as st
import pandas as pd
from io import BytesIO

DEFAULT_DROP_COLUMNS = [
    "Outlet",
    "Supplier Name",
    "Reference PO ID",
    "External Order ID",
    "Stock Receipts",
    "Invoices",
    "FOB",
    "Internal Comments",
    "Original ETD",
    "ETA First Receipt Date",
    "Last Receipt Date",
    "Closed",
    "Sailed",
    "Ship Status",
    "Deposit",
    "Dep Due",
    "Bal Due",
    "Prod Type",
    "Season",
    "Product ID",
    "Supplier SKU",
    "Supplier SKU 2",
    "Manufacturer SKU",
    "Disabled (True/False)",
    "Bin",
    "Special Order Qty",
    "Received Qty",
    "Cancelled Qty",
    "Remaining Qty",
    "Back Order Qty",
    "Tot Buy Ex",
    "Tot COGS Ex",
    "Tot Invoiced Value Ex",
    "POS Price (Ex)",
    "Cust Back Ord Qty",
    "Cust Back Ord COGS",
    "Total Cubic",
    "Created By",
    "Modified By",
    "Modified On",
]

def _norm_col(x: str) -> str:
    # 统一：去首尾空格、合并多空格、小写
    s = str(x).strip().lower()
    s = " ".join(s.split())
    return s

def render(read_excel_any):
    st.subheader("🧹 Excel Column Remover")
    st.markdown("上传 Excel，自动删除你指定的列，并下载清理后的文件。")

    file = st.file_uploader("Upload Excel", type=["xlsx", "xls"])
    if not file:
        st.info("请先上传一个 Excel 文件。")
        return

    raw_bytes = file.getvalue()

    # 读取
    df = read_excel_any(BytesIO(raw_bytes))
    st.caption(f"Rows: {len(df)} | Cols: {len(df.columns)}")

    with st.expander("📋 Preview & Columns", expanded=False):
        st.dataframe(df.head(30), use_container_width=True)
        st.write("Columns:", list(df.columns))

    # 自动匹配默认要删除的列（忽略大小写/空格）
    col_map = {_norm_col(c): c for c in df.columns}  # norm -> original
    default_found = []
    not_found = []

    for c in DEFAULT_DROP_COLUMNS:
        key = _norm_col(c)
        if key in col_map:
            default_found.append(col_map[key])
        else:
            not_found.append(c)

    st.markdown("### ✅ Columns to remove")
    cols_selected = st.multiselect(
        "选择要删除的列（已自动预选你提供的列）",
        options=list(df.columns),
        default=default_found,
    )

    if not_found:
        with st.expander("⚠️ These default columns were not found in your file (will be ignored)", expanded=False):
            st.write(not_found)

    # 执行删除
    if st.button("🧽 Remove selected columns"):
        cleaned = df.drop(columns=cols_selected, errors="ignore")

        st.success(f"Done! New columns count: {len(cleaned.columns)}")
        st.dataframe(cleaned.head(30), use_container_width=True)

        out = BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            cleaned.to_excel(writer, index=False, sheet_name="Cleaned")
        out.seek(0)

        st.download_button(
            "📥 Download Cleaned Excel",
            data=out,
            file_name="cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
