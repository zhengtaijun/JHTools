# 原始导入保持不变
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from rapidfuzz import process, fuzz
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from io import BytesIO
import requests
from PIL import Image
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

# ========== GLOBAL CONFIG ==========
favicon = Image.open("favicon.png")
st.set_page_config(
    page_title="JHCH Tools Suite | Andy Wang",
    layout="centered",
    page_icon=favicon
)
st.title("🛠️ Jory Henley CHC – Internal Tools Suite")
st.caption("© 2025 • App author: **Andy Wang**")

# ========== SIDEBAR NAVIGATION ==========
tool = st.sidebar.radio(
    "🧰 Select a tool:",
    ["TRF Volume Calculator", "Order Merge Tool", "Profit Calculator", "List Split", "Image Table Extractor", "Google Sheet Query"],
    index=0
)

# ========== TOOL 1: TRF Volume Calculator ==========
if tool == "TRF Volume Calculator":
    st.subheader("📦 TRF Volume Calculator")
    st.markdown("📺 [Instructional video](https://youtu.be/S10a3kPEXZg)")

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
        df = pd.concat([
            df,
            pd.DataFrame({"Total Volume": [df["Total Volume"].sum()]})
        ], ignore_index=True)
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
    pass  # 此处省略原代码，完整保留在你现有项目中

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
    pass
# ========== TOOL 3: Profit Calculator ==========
elif tool == "Profit Calculator":
    st.subheader("💰 Profit Calculator")
    st.caption("All data is calculated locally · Multi-product supported · Created by Andy Wang")

    # 1. Product Cost
    cost_mode = st.radio(
        "Input mode:",
        ["Total order cost", "Individual product cost"],
        horizontal=True
    )
    costs = []
    if cost_mode == "Total order cost":
        total_cost = st.number_input(
            "Total order cost (NZD)",
            min_value=0.0,
            step=0.01,
            format="%.2f",
            value=None,
            placeholder="E.g. 350.00"
        ) or 0.0
    else:
        num_products = st.number_input(
            "Number of products",
            min_value=1,
            max_value=20,
            value=2
        )
        cols = st.columns(num_products)
        for i in range(num_products):
            c = cols[i].number_input(
                f"Product {i+1} cost",
                min_value=0.0,
                step=0.01,
                format="%.2f",
                value=None,
                placeholder="E.g. 289.75",
                key=f"cost_{i}"
            ) or 0.0
            costs.append(c)
        total_cost = sum(costs)

    # 2. Product Volume (m³)
    vol_mode = st.radio(
        "Input mode:",
        ["Total volume", "Individual product volume"],
        horizontal=True
    )
    volumes = []
    if vol_mode == "Total volume":
        total_volume = st.number_input(
            "Total volume (m³)",
            min_value=0.0,
            step=0.0001,
            format="%.3f",
            value=None,
            placeholder="E.g. 0.75"
        ) or 0.0
    else:
        num_vols = num_products if cost_mode == "Individual product cost" else st.number_input(
            "Number of volume products",
            min_value=1,
            max_value=20,
            value=2
        )
        cols_v = st.columns(num_vols)
        for i in range(num_vols):
            v = cols_v[i].number_input(
                f"Product {i+1} volume",
                min_value=0.0,
                step=0.0001,
                format="%.3f",
                value=None,
                placeholder="E.g. 0.15",
                key=f"volume_{i}"
            ) or 0.0
            volumes.append(v)
        total_volume = sum(volumes)

    # 3. Shipping Unit Price
    shipping_unit_price = st.number_input(
        "Shipping unit price (NZD/m³, GST not included, default 150)",
        min_value=0.0,
        step=0.01,
        format="%.2f",
        value=150.0
    )

    # 4. Sale Price
    sale_price = st.number_input(
        "Input sale price (GST included, NZD)",
        min_value=0.0,
        step=0.01,
        format="%.2f",
        value=None,
        placeholder="E.g. 1200"
    ) or 0.0

    # Calculations
    gst_cost = total_cost * 1.15
    shipping_cost = total_volume * shipping_unit_price
    shipping_gst = shipping_cost * 1.15
    rent = sale_price * 0.10
    jcd = sale_price * 0.09
    total_expense = gst_cost + shipping_gst + rent + jcd
    profit_with_gst = sale_price - total_expense
    profit_no_gst = profit_with_gst / 1.15 if profit_with_gst else 0.0

    def pct(n):
        return f"{(n/(sale_price or 1)*100):.2f}" + "%" if sale_price else "-"

    result_rows = [
        ["COGS", gst_cost, pct(gst_cost)],
        ["Shipping", shipping_gst, pct(shipping_gst)],
        ["Rent (10%)", rent, pct(rent)],
        ["JCD Cost (9%)", jcd, pct(jcd)],
        ["Total Cost", total_expense, pct(total_expense)],
        ["Profit (incl. GST)", profit_with_gst, pct(profit_with_gst)],
        ["Profit (excl. GST)", profit_no_gst, ""]
    ]
    df_res = pd.DataFrame(result_rows, columns=["Item", "Amount (NZD)", "Ratio to Sale Price"])
    df_res["Amount (NZD)"] = df_res["Amount (NZD)"].map(lambda x: f"{x:.2f}")
    st.table(df_res)

    with st.expander("Calculation details"):
        st.markdown(f"""
- **Total order cost** = {total_cost:.2f} NZD
- **COGS** = Total order cost × 1.15 = {gst_cost:.2f} NZD
- **Total volume** = {total_volume:.3f} m³
- **Shipping (no GST)** = Total volume × Shipping unit price = {shipping_cost:.2f} NZD
- **Shipping (GST included)** = Shipping × 1.15 = {shipping_gst:.2f} NZD
- **COGS & Shipping** = COGS + Shipping = {gst_cost+shipping_gst:.2f} NZD
- **Rent** = Sale price × 10% = {rent:.2f} NZD
- **JCD Cost** = Sale price × 9% = {jcd:.2f} NZD
- **Total cost** = COGS & Shipping + Rent + JCD Cost = {total_expense:.2f} NZD
- **Profit (incl. GST)** = Sale price - Total cost = {profit_with_gst:.2f} NZD
- **Profit (excl. GST)** = Profit (incl. GST) / 1.15 = {profit_no_gst:.2f} NZD
        """
    )

    if sale_price > 0:
        fig, ax = plt.subplots()
        ax.pie(
            [gst_cost, shipping_gst, rent, jcd, max(profit_with_gst, 0)],
            labels=["COGS", "Shipping", "Rent", "JCD", "Profit"],
            autopct="%1.1f%%",
            startangle=90
        )
        ax.axis("equal")
        st.pyplot(fig)

    if st.button("Export results to Excel"):
        out = BytesIO()
        df_res.to_excel(out, index=False)
        out.seek(0)
        st.download_button(
            "📥 Download Excel",
            out,
            file_name="profit_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    pass
    pass
# ========== TOOL 4: List Split ==========
# ========== TOOL 4: List Split ==========
elif tool == "List Split":
    st.subheader("📄 List Split")
    st.markdown("Paste copied table data with order number and products. Format: `2*Chair,1*Table`")

    pasted_text = st.text_area("Paste your copied data below (from Excel):")

    if st.button("🔍 Analyze pasted content") and pasted_text:
        try:
            from io import StringIO
            df_input = pd.read_csv(StringIO(pasted_text), sep="\t", header=None)

            st.write("✅ Preview of parsed input:")
            st.dataframe(df_input)

            records = []
            for _, row in df_input.iterrows():
                order_id = str(row.iloc[0])
                product_str = str(row.iloc[-1])
                items = [item.strip() for item in product_str.split(',') if '*' in item]

                for item in items:
                    try:
                        qty_str, name = item.split('*', 1)
                        qty = int(qty_str.strip())
                        name = name.strip()
                        records.append({
                            'name': name,
                            'order': order_id,
                            'qty': qty
                        })
                    except ValueError:
                        st.warning(f"⚠️ Skipped malformed item: {item}")

            if records:
                df_result = pd.DataFrame(records)
                st.success("✅ Processing completed.")
                st.dataframe(df_result)

                to_download = BytesIO()
                df_result.to_excel(to_download, index=False)
                to_download.seek(0)

                st.download_button("📥 Download Excel", to_download, file_name="parsed_list.xlsx")
            else:
                st.error("No valid records found. Please check your input.")
        except Exception as e:
            st.error(f"❌ Error processing input: {e}")
            pass
# ========== TOOL 5: Image Table Extractor ==========
elif tool == "Image Table Extractor":
    st.subheader("🖼️ Excel Screenshot to Table")
    st.markdown("Paste (Ctrl+V) or drag a screenshot of an Excel table. Supported formats: JPG, PNG")

    from PIL import Image
    import pytesseract
    import re
    import base64
    import json
    import streamlit.components.v1 as components

    uploaded_image = st.file_uploader("Upload Screenshot", type=["jpg", "jpeg", "png"])

    pasted_image_bytes = st.session_state.get("pasted_image", None)

    # Paste listener
    pasted_image_holder = st.empty()
    components.html('''
        <script>
        document.addEventListener('paste', async function (event) {
            const items = (event.clipboardData || window.clipboardData).items;
            for (const item of items) {
                if (item.type.indexOf('image') === 0) {
                    const file = item.getAsFile();
                    const reader = new FileReader();
                    reader.onload = function(event) {
                        const base64Image = event.target.result;
                        const pyMsg = {type: "paste_image", data: base64Image};
                        window.parent.postMessage(pyMsg, "*");
                    };
                    reader.readAsDataURL(file);
                }
            }
        });
        </script>
    ''', height=0)

    # Paste upload handler
    pasted_json = st.query_params.get("pasted_image")
    if pasted_json:
        try:
            imgdata = base64.b64decode(pasted_json[0].split(",")[-1])
            st.session_state.pasted_image = imgdata
            st.query_params.clear()  # Clear param after use
        except:
            st.warning("Failed to decode pasted image.")

    image = None
    if uploaded_image:
        image = Image.open(uploaded_image)
        st.image(image, caption="Uploaded image", use_column_width=True)
    elif pasted_image_bytes:
        image = Image.open(BytesIO(pasted_image_bytes))
        st.image(image, caption="Pasted image", use_column_width=True)

    if image:
        with st.spinner("Running OCR... please wait..."):
            raw_text = pytesseract.image_to_string(image)
            lines = raw_text.strip().split("\n")
            rows = [re.split(r'\t+|\s{2,}', line.strip()) for line in lines if line.strip()]
            max_len = max((len(row) for row in rows), default=0)
            rows = [row + [''] * (max_len - len(row)) for row in rows]
            df = pd.DataFrame(rows)

        st.success("✅ OCR complete. Here's the extracted table:")
        st.dataframe(df)

        # Download Excel
        out = BytesIO()
        df.to_excel(out, index=False, header=False)
        out.seek(0)
        st.download_button(
            "📥 Download as Excel",
            out,
            file_name="extracted_table.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Copy to clipboard button (tab-separated)
        def df_to_tsv(df):
            return '\n'.join(['\t'.join(map(str, row)) for row in df.values.tolist()])

        tsv_string = df_to_tsv(df)

        components.html(f'''
            <textarea id="tsv" style="position:absolute;left:-1000px">{tsv_string}</textarea>
            <button onclick="copyTSV()">📋 Copy Table (for Excel/Sheets)</button>
            <script>
            function copyTSV() {{
                const t = document.getElementById("tsv");
                t.select();
                document.execCommand("copy");
                alert("✅ Table copied to clipboard. You can now paste into Excel or Google Sheets.");
            }}
            </script>
        ''', height=50)
    else:
        st.info("Please upload or paste a screenshot of a table to begin.")
        pass
# ========== TOOL 6: Order Check ==========
elif tool == "Google Sheet Query":
    st.subheader("🔎 Google Sheet 查询工具")
    st.markdown("使用 Google Sheet 作为数据库，固定提取第 1、2、4、6、7、13、15 列")

    SHEET_ID = "17twAYxaakAIbDhQvFR6FdgUVFHxBJlL5w8rrC7gpCu8"
    SHEET_NAME = "Sheet1"

    @st.cache_data
    def load_sheet_data():
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)

        all_data = sheet.get_all_values()
        if not all_data:
            return pd.DataFrame()

        # 固定选择第1, 2, 4, 6, 7, 13, 15列 (注意 Python索引从0开始)
        col_indices = [0, 1, 3, 5, 6, 12, 14]  # A, B, D, F, G, M, O
        headers = all_data[0]
        rows = all_data[1:]

        # 处理：只保留指定列
        selected_headers = [headers[i] if i < len(headers) else f"Col{i+1}" for i in col_indices]
        selected_rows = [[row[i] if i < len(row) else "" for i in col_indices] for row in rows]

        df = pd.DataFrame(selected_rows, columns=selected_headers)
        return df

    try:
        df = load_sheet_data()
        if df.empty:
            st.warning("⚠️ 表格为空或数据加载失败。")
        else:
            st.success("✅ 表格加载成功！")
            
            with st.expander("📋 显示全部数据（可选）", expanded=False):
                st.dataframe(df, use_container_width=True)

            query = st.text_input("🔍 输入关键词（模糊匹配所有列）:")

            if query:
                filtered = df[df.apply(lambda row: row.astype(str).str.contains(query, case=False).any(), axis=1)]
                st.markdown(f"🔎 **共找到 {len(filtered)} 条匹配结果：**")
                st.dataframe(filtered, use_container_width=True)
            else:
                st.info("请输入关键词开始查询。")

    except Exception as e:
        st.error(f"❌ 加载 Google Sheet 失败：{e}")




