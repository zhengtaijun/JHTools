from io import BytesIO, StringIO
import pandas as pd
import streamlit as st


def render():
    st.subheader("📄 List Split")
    st.markdown("Paste copied table data with order number and products. Format: `2*Chair,1*Table`")

    pasted_text = st.text_area("Paste your copied data below (from Excel):")

    if st.button("🔍 Analyze pasted content") and pasted_text:
        try:
            df_input = pd.read_csv(StringIO(pasted_text), sep="\t", header=None, dtype=str)
            st.write("✅ Preview of parsed input:")
            st.dataframe(df_input, use_container_width=True)

            def _fmt_cell(v):
                if v is None:
                    return ""
                s = str(v).strip()
                return "" if s.lower() in ("nan", "none") else s

            records = []
            for _, row in df_input.iterrows():
                order_id = _fmt_cell(row.iloc[0]) if len(row) >= 1 else ""
                supplier_code = _fmt_cell(row.iloc[-2]) if len(row) >= 2 else ""
                combined_order_ref = f"{supplier_code}//{order_id}" if supplier_code else order_id

                product_str = _fmt_cell(row.iloc[-1]) if len(row) >= 1 else ""
                items = [item.strip() for item in product_str.split(',') if '*' in item]

                for item in items:
                    try:
                        qty_str, name = item.split('*', 1)
                        qty_str = _fmt_cell(qty_str)
                        name = _fmt_cell(name)
                        if not name:
                            continue
                        qty = int(float(qty_str)) if qty_str else 0
                        records.append({"order": combined_order_ref, "name": name, "qty": qty})
                    except Exception:
                        st.warning(f"⚠️ Skipped malformed item: {item}")

            if not records:
                st.error("No valid records found. Please check your input.")
                return

            df_result = pd.DataFrame(records)[["order", "name", "qty"]]
            st.info("🧩 已将倒数第二列识别为『供应商订货号』，并与第一列『订单号』合并为：**供应商订货号//订单号**")
            st.success("✅ Processing completed.")
            st.dataframe(df_result, use_container_width=True)

            to_download = BytesIO()
            df_result.to_excel(to_download, index=False)
            to_download.seek(0)

            st.download_button("📥 Download Excel", to_download, file_name="parsed_list.xlsx")
        except Exception as e:
            st.error(f"❌ Error processing input: {e}")
