# tools/excel_column_remover.py
import streamlit as st
import pandas as pd
from io import BytesIO

DEFAULT_DROP_COLUMNS = [
    "Outlet","Supplier Name","Reference PO ID","External Order ID","Stock Receipts","Invoices",
    "FOB","Internal Comments","Original ETD","ETA","First Receipt Date","Last Receipt Date","Closed",
    "Sailed","Ship Status","Deposit","Dep Due","Bal Due","Prod Type","Season","Product ID",
    "Supplier SKU","Supplier SKU 2","Manufacturer SKU","Disabled (True/False)","Bin",
    "Special Order Qty","Received Qty","Cancelled Qty","Remaining Qty","Back Order Qty",
    "Tot Buy Ex","Tot COGS Ex","Tot Invoiced Value Ex","POS Price (Ex)","Cust Back Ord Qty",
    "Cust Back Ord COGS","Total Cubic","Created By","Modified By","Modified On"
]

def _norm_col(x: str) -> str:
    s = str(x).strip().lower()
    s = " ".join(s.split())
    return s

def _read_excel(file_bytes: bytes) -> pd.DataFrame:
    """
    优先用你项目里的 read_excel_any（如果存在），否则用 pandas 默认读。
    """
    try:
        # 你如果已经把 read_excel_any 放到 utils 里，按你的实际文件名改这里
        # 常见：from utils.excel_loader import read_excel_any
        from utils.excel_loader import read_excel_any  # <- 如果你文件名不是这个，改成你自己的
        return read_excel_any(BytesIO(file_bytes))
    except Exception:
        # 兜底：普通读取
        return pd.read_excel(BytesIO(file_bytes))

def render():
    st.subheader("🧹 Excel Column Remover")
    st.markdown("上传 Excel，自动删除指定列并下载清理后的文件。")

    file = st.file_uploader("Upload Excel", type=["xlsx", "xls"])
    if not file:
        st.info("请先上传一个 Excel 文件。")
        return

    raw_bytes = file.getvalue()

    try:
        df = _read_excel(raw_bytes)
    except Exception as e:
        st.error(f"❌ Failed to read Excel: {e}")
        return

    st.caption(f"Rows: {len(df)} | Cols: {len(df.columns)}")

    with st.expander("📋 Preview & Columns", expanded=False):
        st.dataframe(df.head(30), use_container_width=True)
        st.write("Columns:", list(df.columns))

    # 默认预选（忽略大小写/空格）
    col_map = {_norm_col(c): c for c in df.columns}
    default_found, not_found = [], []

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
        with st.expander("⚠️ These default columns were not found (ignored)", expanded=False):
            st.write(not_found)

    if st.button("🧽 Remove selected columns"):
        cleaned = df.drop(columns=cols_selected, errors="ignore")

        st.success(f"✅ Done! New columns count: {len(cleaned.columns)}")
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

