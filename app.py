import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz
from concurrent.futures import ThreadPoolExecutor
from datetime import timedelta, datetime
from io import BytesIO
import requests
from PIL import Image
import streamlit as st

favicon = Image.open("favicon.png")

st.set_page_config(
    page_title="JHCH Tools",
    page_icon=favicon,  # ← 设置 favicon
    layout="centered"
)

# ────────────────────────────────────────────────────────────────
# GLOBAL PAGE CONFIG
# ────────────────────────────────────────────────────────────────
st.set_page_config(page_title="JHCH Tools Suite | Andy Wang", layout="centered")
st.title("🛠️ Jory Henley CHC – Internal Tools Suite")
st.caption("© 2025 • App author: **Andy Wang**")

# ────────────────────────────────────────────────────────────────
# SIDEBAR – TOOL SELECTOR
# ────────────────────────────────────────────────────────────────
tool = st.sidebar.radio(
    "🧰 Select a tool:",
    ["TRF Volume Calculator", "Order Merge Tool"],
    index=0,
    
)

# =================================================================
# 1) TRF VOLUME CALCULATOR
# =================================================================
if tool == "TRF Volume Calculator":

    st.subheader("📦 TRF Volume Calculator")
    st.markdown(
        "📺 **Need help?** Watch the "
        "[instructional video here](https://youtu.be/S10a3kPEXZg)"
    )

    # ---------- CONFIG ----------
    PRODUCT_INFO_URL = (
        "https://raw.githubusercontent.com/zhengtaijun/JHCH_TRF-Volume/main/product_info.xlsx"
    )

    @st.cache_data
    def load_product_info():
        response = requests.get(PRODUCT_INFO_URL)
        response.raise_for_status()
        df = pd.read_excel(BytesIO(response.content))
        with st.expander("✅ Product-info file loaded. Click to view columns", expanded=False):
            st.write(df.columns.tolist()) 
        if {"Product Name", "CBM"} - set(df.columns):
            raise ValueError("`Product Name` and `CBM` columns are required.")
        names = df["Product Name"].fillna("").astype(str)
        cbms = pd.to_numeric(df["CBM"], errors="coerce").fillna(0)
        return dict(zip(names.tolist(), cbms.tolist())), names.tolist()

    product_dict, product_names_all = load_product_info()

    # ---------- UI ----------
    warehouse_file = st.file_uploader("Upload warehouse export (Excel)", type=["xlsx"])
    col_prod = st.number_input("Column # of **Product Name**", min_value=1, value=3)
    col_order = st.number_input("Column # of **Order Number**", min_value=1, value=7)
    col_qty = st.number_input("Column # of **Quantity**", min_value=1, value=8)

    # ---------- Helpers ----------
    def match_product(name: str):
        if name in product_dict:
            return product_dict[name]
        match, score, _ = process.extractOne(name, product_names_all, scorer=fuzz.partial_ratio)
        return product_dict[match] if score >= 80 else None

    def process_volume_file(file, p_col, q_col):
        df = pd.read_excel(file)
        st.write("📊 Warehouse file loaded. Shape:", df.shape)

        product_names = df.iloc[:, p_col].fillna("").astype(str).tolist()
        quantities = pd.to_numeric(df.iloc[:, q_col], errors="coerce").fillna(0)

        st.write("🧾 Sample product names:", product_names[:5])
        st.write("🔢 Sample quantities:", quantities.head().tolist())

        total = len(product_names)
        volumes: list[float | None] = []

        def worker(start: int, end: int):
            partial = []
            for i in range(start, end):
                name = product_names[i].strip()
                vol = match_product(name) if name else None
                partial.append(vol)
            return partial

        with ThreadPoolExecutor(max_workers=4) as pool:
            chunk = max(total // 4, 1)
            futures = [
                pool.submit(worker, i * chunk, (i + 1) * chunk if i < 3 else total)
                for i in range(4)
            ]
            for f in futures:
                volumes.extend(f.result())

        st.write("🧮 Volume matching done. First 10:", volumes[:10])
        df["Volume"] = pd.to_numeric(pd.Series(volumes), errors="coerce").fillna(0)
        df["Total Volume"] = df["Volume"] * quantities
        st.write("✅ Columns ‘Volume’ and ‘Total Volume’ added")

        df = pd.concat([df, pd.DataFrame({"Total Volume": [df["Total Volume"].sum()]})], ignore_index=True)
        return df

    # ---------- Run ----------
    if warehouse_file and st.button("Calculate volume"):
        with st.spinner("Processing…"):
            try:
                result_df = process_volume_file(warehouse_file, col_prod - 1, col_qty - 1)
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    result_df.to_excel(writer, index=False)
                buffer.seek(0)
                stamp = datetime.now().strftime("%Y%m%d%H%M%S")
                st.success("✅ Done. Download below:")
                st.download_button(
                    "📥 Download Excel",
                    data=buffer,
                    file_name=f"TRF_Volume_Result_{stamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"❌ Error: {e}")


# =================================================================
# 2) ORDER MERGE TOOL
# =================================================================
else:  # tool == "Order Merge Tool"

    st.subheader("📋 Order Merge Tool")
    st.sidebar.markdown("📘 [View User Guide](https://github.com/zhengtaijun/JHTools/blob/main/docs/instructions.md)")


    # ---------- Upload ----------
    file1 = st.file_uploader("Upload File 1", type=["xlsx"], key="merge1")
    file2 = st.file_uploader("Upload File 2", type=["xlsx"], key="merge2")

    # ---------- Helper funcs ----------
    def clean_phone(num):
        if pd.notna(num):
            num = str(int(num)) if isinstance(num, float) else str(num)
            return num.strip()
        return ""

    def target_wednesday(inv_date):
        if not pd.isna(inv_date) and not isinstance(inv_date, pd.Timestamp):
            inv_date = pd.to_datetime(inv_date)
        weekday = inv_date.weekday()
        days = 2 - weekday if weekday <= 2 else 9 - weekday
        return (inv_date + timedelta(days=days)).date()

    def process_merge(f1, f2):
        df1, df2 = pd.read_excel(f1), pd.read_excel(f2)
        has1, has2 = "Freight Ex" in df1.columns, "Freight Ex" in df2.columns
        if has1 and has2:
            st.error("Both files contain **Freight Ex**; only one should.")
            return None
        if not has1 and not has2:
            st.error("Neither file contains **Freight Ex**.")
            return None
        df_freight = df1 if has1 else df2
        df_main = df2 if has1 else df1

        freight_map = dict(zip(df_freight["Order No"], df_freight["Freight Ex"]))
        rows = []

        for order_no, grp in df_main.groupby("Order No"):
            inv_date = grp["Inv Date"].iloc[0]
            row = [
                order_no,
                inv_date,
                target_wednesday(inv_date),
                1 if freight_map.get(order_no, 0) > 0 else "pickup",
                "",
                grp["Bill Name"].iloc[0],
                " ".join(filter(None, [clean_phone(grp["Billing Phone"].iloc[0]),
                                       clean_phone(grp["Billing Mobile"].iloc[0])])),
                "", "", "", "",
                1 if grp["Order Status"].iloc[0] == "Awaiting Payment" else "",
                "", "",
                ",".join(f"{int(r['Item Qty'])}*{r['Short Description']}" for _, r in grp.iterrows())
            ]
            rows.append(row)

        return pd.DataFrame(rows)

    # ---------- Run ----------
    if file1 and file2 and st.button("Merge orders"):
        with st.spinner("Processing…"):
            try:
                merged = process_merge(file1, file2)
                if merged is not None:
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
                        merged.to_excel(writer, index=False, header=False)
                    out.seek(0)
                    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
                    st.success("✅ Merge complete. Download below:")
                    st.download_button(
                        "📥 Download Merged Excel",
                        data=out,
                        file_name=f"order_merge_{stamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            except Exception as e:
                st.error(f"❌ Error: {e}")
