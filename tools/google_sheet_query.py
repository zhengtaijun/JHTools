import json
import pandas as pd
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials

from utils.constants import SHEET_ID, SHEET_NAME


def render():
    st.subheader("🔎 Google Sheet 查询工具")
    st.markdown("使用 Google Sheet 作为数据库，固定提取第 1、2、4、6、7、13、15 列")

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

        col_indices = [0, 1, 3, 5, 6, 12, 14]  # A, B, D, F, G, M, O
        headers = all_data[0]
        rows = all_data[1:]

        selected_headers = [headers[i] if i < len(headers) else f"Col{i+1}" for i in col_indices]
        selected_rows = [[row[i] if i < len(row) else "" for i in col_indices] for row in rows]
        return pd.DataFrame(selected_rows, columns=selected_headers)

    try:
        df = load_sheet_data()
        if df.empty:
            st.warning("⚠️ 表格为空或数据加载失败。")
            return

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
