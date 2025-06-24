# ✅ This is the updated `app.py` with 3 tools integrated

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from rapidfuzz import process, fuzz
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from io import BytesIO
import requests
from PIL import Image

# ========== GLOBAL CONFIG ==========
st.set_page_config(
    page_title="JHCH Tools Suite | Andy Wang",
    layout="centered",
    page_icon=Image.open("favicon.png")
)
st.title("🛠️ Jory Henley CHC – Internal Tools Suite")
st.caption("© 2025 • App author: **Andy Wang**")

# ========== SIDEBAR NAVIGATION ==========
tool = st.sidebar.radio(
    "🧰 Select a tool:",
    ["TRF Volume Calculator", "Order Merge Tool", "Profit Calculator"],
    index=0
)

# ========== TOOL 1: TRF Volume Calculator ==========
if tool == "TRF Volume Calculator":
    st.subheader("📦 TRF Volume Calculator")
    st.markdown("📺 [Instructional video](https://youtu.be/S10a3kPEXZg)")

    PRODUCT_INFO_URL = "https://raw.githubusercontent.com/zhengtaijun/JHCH_TRF-Volume/main/product_info.xlsx"

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

    warehouse_file = st.file_uploader("Upload warehouse export (Excel)", type=["xlsx"])
    col_prod = st.number_input("Column # of **Product Name**", min_value=1, value=3)
    col_order = st.number_input("Column # of **Order Number**", min_value=1, value=7)
    col_qty = st.number_input("Column # of **Quantity**", min_value=1, value=8)

    def match_product(name: str):
        if name in product_dict:
            return product_dict[name]
        match, score, _ = process.extractOne(name, product_names_all, scorer=fuzz.partial_ratio)
        return product_dict[match] if score >= 80 else None

    def process_volume_file(file, p_col, q_col):
        df = pd.read_excel(file)
        product_names = df.iloc[:, p_col].fillna("").astype(str).tolist()
        quantities = pd.to_numeric(df.iloc[:, q_col], errors="coerce").fillna(0)

        total = len(product_names)
        volumes = []

        def worker(start: int, end: int):
            partial = []
            for i in range(start, end):
                name = product_names[i].strip()
                vol = match_product(name) if name else None
                partial.append(vol)
            return partial

        with ThreadPoolExecutor(max_workers=4) as pool:
            chunk = max(total // 4, 1)
            futures = [pool.submit(worker, i * chunk, (i + 1) * chunk if i < 3 else total) for i in range(4)]
            for f in futures:
                volumes.extend(f.result())

        df["Volume"] = pd.to_numeric(pd.Series(volumes), errors="coerce").fillna(0)
        df["Total Volume"] = df["Volume"] * quantities
        df = pd.concat([df, pd.DataFrame({"Total Volume": [df["Total Volume"].sum()]})], ignore_index=True)
        return df

    if warehouse_file and st.button("Calculate volume"):
        with st.spinner("Processing…"):
            try:
                result_df = process_volume_file(warehouse_file, col_prod - 1, col_qty - 1)
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    result_df.to_excel(writer, index=False)
                buffer.seek(0)
                st.download_button("📥 Download Excel", buffer, file_name="TRF_Volume_Result.xlsx")
            except Exception as e:
                st.error(f"❌ Error: {e}")

# ========== TOOL 2: Order Merge Tool ==========
elif tool == "Order Merge Tool":
    st.subheader("📋 Order Merge Tool")
    st.markdown("📘 [View User Guide](https://github.com/zhengtaijun/JHTools/blob/main/instructions.md)")

    file1 = st.file_uploader("Upload File 1", type=["xlsx"], key="merge1")
    file2 = st.file_uploader("Upload File 2", type=["xlsx"], key="merge2")

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
                " ".join(filter(None, [clean_phone(grp["Billing Phone"].iloc[0]), clean_phone(grp["Billing Mobile"].iloc[0])])),
                "", "", "", "",
                1 if grp["Order Status"].iloc[0] == "Awaiting Payment" else "",
                "", "",
                ",".join(f"{int(r['Item Qty'])}*{r['Short Description']}" for _, r in grp.iterrows())
            ]
            rows.append(row)

        return pd.DataFrame(rows)

    if file1 and file2 and st.button("Merge orders"):
        with st.spinner("Processing…"):
            try:
                merged = process_merge(file1, file2)
                if merged is not None:
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
                        merged.to_excel(writer, index=False, header=False)
                    out.seek(0)
                    st.download_button("📥 Download Merged Excel", out, file_name="order_merge.xlsx")
            except Exception as e:
                st.error(f"❌ Error: {e}")

# ========== TOOL 3: Profit Calculator ==========
else:
    st.subheader("💰 Profit Calculator")
    cost_mode = st.radio("1. Product Cost Input Mode:", ["Total", "Individual"], horizontal=True)
    costs = []
    if cost_mode == "Total":
        total_cost = st.number_input("Total order cost (NZD)", min_value=0.0, step=0.01, value=None)
    else:
        num = st.number_input("Number of products", min_value=1, max_value=20, value=2)
        cols = st.columns(int(num))
        for i in range(int(num)):
            c = cols[i].number_input(f"Product {i+1} cost", min_value=0.0, step=0.01, value=None, key=f"c{i}")
            costs.append(c or 0)
        total_cost = sum(costs)

    vol_mode = st.radio("2. Volume Input Mode:", ["Total", "Individual"], horizontal=True)
    vols = []
    if vol_mode == "Total":
        total_volume = st.number_input("Total volume (m³)", min_value=0.0, step=0.001, value=None)
    else:
        count = int(num) if cost_mode == "Individual" else st.number_input("Number of volume products", 1, 20, 2)
        vcols = st.columns(count)
        for i in range(count):
            v = vcols[i].number_input(f"Product {i+1} volume", min_value=0.0, step=0.001, value=None, key=f"v{i}")
            vols.append(v or 0)
        total_volume = sum(vols)

    ship_unit = st.number_input("3. Shipping Price per m³ (excl. GST)", min_value=0.0, step=0.01, value=150.0)
    sale_price = st.number_input("4. Sale Price (incl. GST)", min_value=0.0, step=0.01, value=None)

    gst_cost = (total_cost or 0) * 1.15
    shipping_cost = (total_volume or 0) * (ship_unit or 0) * 1.15
    rent = (sale_price or 0) * 0.10
    jcd = (sale_price or 0) * 0.09
    total_expense = gst_cost + shipping_cost + rent + jcd
    profit = (sale_price or 0) - total_expense
    profit_ex_gst = profit / 1.15 if profit else 0

    def pct(n): return f"{(n/(sale_price or 1)*100):.1f}%" if sale_price else "-"

    df = pd.DataFrame([
        ["COGS", gst_cost, pct(gst_cost)],
        ["Shipping", shipping_cost, pct(shipping_cost)],
        ["Rent", rent, pct(rent)],
        ["JCD Cost", jcd, pct(jcd)],
        ["Total Cost", total_expense, pct(total_expense)],
        ["Profit (incl. GST)", profit, pct(profit)],
        ["Profit (excl. GST)", profit_ex_gst, ""]
    ], columns=["Item", "Amount (NZD)", "Ratio to Sale"])
    df["Amount (NZD)"] = df["Amount (NZD)"].map(lambda x: f"{x:.2f}")
    st.table(df)

    if sale_price:
        fig, ax = plt.subplots()
        ax.pie([gst_cost, shipping_cost, rent, jcd, max(profit, 0)],
               labels=["COGS", "Shipping", "Rent", "JCD", "Profit"],
               autopct="%1.1f%%", startangle=90)
        ax.axis("equal")
        st.pyplot(fig)

    if st.button("Export to Excel"):
        out = BytesIO()
        df.to_excel(out, index=False)
        out.seek(0)
        st.download_button("📥 Download Excel", out, file_name="profit_result.xlsx")
