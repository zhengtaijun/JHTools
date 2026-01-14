import re
from io import BytesIO

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from utils.excel_io import read_excel_any


def render():
    st.subheader("📋 Order Merge Tool V2")
    st.markdown("📘 [View User Guide](https://github.com/zhengtaijun/JHTools/blob/main/instruction%20v2.png)")

    st.info(
        "📢 公告：本工具将旧表（按产品分行）整理为每个 **OrderNumber** 只保留一行的新表。\n\n"
        "- 产品名自动去重合并（Product_Description + Size + Colour）\n"
        "- 第14列输出 PO（形如 PO3513），忽略 “2 Available”\n"
        "- 第15列输出 Items（形如 `qty*合并后产品名`，多件逗号分隔）\n"
        "- **第二列 DateCreated 输出为 `yyyy/mm/dd`（按斜杠位置暴力重排）**\n"
        "- 预览支持**一键复制表格（不含表头）**\n"
    )

    file = st.file_uploader("Upload the Excel file (old layout)", type=["xlsx", "xls"], key="order_merge_v2")

    RE_PO = re.compile(r'(?:PO:|<strong>PO:</strong>)\s*#?\s*(\d+)', re.IGNORECASE)
    RE_WS = re.compile(r'\s+')

    REQUIRED_COLS = [
        "DateCreated", "OrderNumber", "OrderStatus", "Product_Description", "Size", "Colour",
        "CustomerName", "Phone", "Mobile", "DeliveryMode", "PublicComments", "qtyRequired", "SourceFrom"
    ]

    RE_DMY = re.compile(r'(\d{1,4})\s*/\s*(\d{1,2})\s*/\s*(\d{2,4})')
    RE_YMD_FINAL = re.compile(r'^\s*\d{4}/\d{1,2}/\d{1,2}\s*$')

    def brutal_extract_ymd(value):
        s = str(value).strip()
        m = RE_DMY.search(s)
        if not m:
            return None
        a, b, c = m.groups()
        day = int(a)
        month = int(b)
        year = int(c) if len(c) == 4 else int("20" + c)
        return (year, month, day)

    def brutal_format_ymd(value):
        s = str(value).strip()
        if RE_YMD_FINAL.match(s):
            return s
        t = brutal_extract_ymd(s)
        if not t:
            return s if value is not None else ""
        y, m, d = t
        return f"{y}/{m}/{d}"

    def brutal_min_date(series):
        tuples = [brutal_extract_ymd(v) for v in series]
        tuples = [t for t in tuples if t is not None]
        if not tuples:
            for v in series:
                if str(v).strip():
                    return brutal_format_ymd(v)
            return ""
        y, m, d = min(tuples)
        return f"{y}/{m}/{d}"

    def clean_str(s):
        if pd.isna(s):
            return ""
        s = str(s)
        s = re.sub(r"<[^>]*>", "", s)
        s = RE_WS.sub(" ", s).strip()
        return s

    def contains_ci(hay, needle):
        return bool(needle) and needle.lower() in hay.lower()

    def merge_product_name(prod, size, colour):
        p = clean_str(prod)
        s = clean_str(size)
        c = clean_str(colour)

        parts = [p] if p else []
        if s and not contains_ci(p, s):
            parts.append(s)
        merged = " ".join(parts) if parts else ""
        if c and not contains_ci(merged, c):
            merged = (merged + (" - " if merged else "") + c)
        return merged

    def fmt_qty_name(qty, name):
        if not name:
            return ""
        if pd.isna(qty) or str(qty).strip() == "":
            return name
        try:
            q = float(qty)
            q_int = int(q)
            q_str = str(q_int) if abs(q - q_int) < 1e-9 else str(q)
        except Exception:
            q_str = str(qty).strip()
        return f"{q_str}*{name}"

    def extract_pos(value):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return []
        text = str(value)
        if re.search(r"\b\d+\s+Available\b", text, flags=re.IGNORECASE):
            return []
        matches = RE_PO.findall(text)
        return [f"PO{m}" for m in matches]

    def first_nonempty(values):
        for v in values:
            if isinstance(v, str) and v.strip():
                return v.strip()
            if pd.notna(v) and str(v).strip():
                return str(v).strip()
        return ""

    def consolidate(df: pd.DataFrame) -> pd.DataFrame:
        for col in REQUIRED_COLS:
            if col not in df.columns:
                df[col] = pd.NA

        df["_MergedName"] = df.apply(
            lambda r: merge_product_name(r["Product_Description"], r["Size"], r["Colour"]), axis=1
        )
        df["_ItemLine"] = [fmt_qty_name(q, n) for q, n in zip(df["qtyRequired"], df["_MergedName"])]
        df["_POs"] = df["SourceFrom"].apply(extract_pos)

        def merge_phones(phone, mobile):
            def normalize_phone(v):
                if pd.isna(v):
                    return ""
                s = str(v).strip()
                if s.endswith(".0"):
                    s = s[:-2]
                return s

            parts = [normalize_phone(phone), normalize_phone(mobile)]
            parts = [p for p in parts if p]
            seen, unique = set(), []
            for x in parts:
                if x not in seen:
                    seen.add(x)
                    unique.append(x)
            return ", ".join(unique)

        df["_Phones"] = [merge_phones(p, m) for p, m in zip(df["Phone"], df["Mobile"])]

        rows = []
        for order, g in df.groupby("OrderNumber", dropna=False):
            g = g.copy()

            delivery_vals = [str(x).strip().lower() for x in g["DeliveryMode"].tolist()]
            home_flag = 1 if any(x == "home" for x in delivery_vals) else "pickup"

            status_vals = [str(x).strip() for x in g["OrderStatus"].tolist() if str(x).strip()]
            awaiting_flag = 1 if any(x.lower() == "awaiting payment" for x in status_vals) else ""

            items = [x for x in g["_ItemLine"].tolist() if x]
            items_text = ", ".join(items)

            po_list, seen_po = [], set()
            for sub in g["_POs"].tolist():
                for x in sub:
                    if x not in seen_po:
                        seen_po.add(x)
                        po_list.append(x)
            po_text = ", ".join(po_list)

            phone_opts = [x for x in g["_Phones"].tolist() if x]
            seen_ph, phone_unique = set(), []
            for x in phone_opts:
                if x not in seen_ph:
                    seen_ph.add(x)
                    phone_unique.append(x)
            phones_text = ", ".join(phone_unique)

            comments_vals = [clean_str(x) for x in g["PublicComments"].tolist() if clean_str(x)]
            seen_c, comments_unique = set(), []
            for x in comments_vals:
                if x not in seen_c:
                    seen_c.add(x)
                    comments_unique.append(x)
            comments_text = " | ".join(comments_unique)

            date_value = brutal_min_date(g["DateCreated"])
            customer = first_nonempty(g["CustomerName"].tolist())

            row = {
                "OrderNumber": order,
                "DateCreated": date_value,
                "Col3": "",
                "HomeDelivery": home_flag,
                "Col5": "",
                "CustomerName": customer,
                "ContactPhones": phones_text,
                "Col8": "", "Col9": "", "Col10": "", "Col11": "",
                "AwaitingPayment": awaiting_flag,
                "PublicComments": comments_text,
                "POs": po_text,
                "Items": items_text,
            }
            rows.append(row)

        out = pd.DataFrame(rows)
        out["DateCreated"] = out["DateCreated"].astype(str).apply(brutal_format_ymd)

        out = out[
            ["OrderNumber", "DateCreated", "Col3", "HomeDelivery", "Col5",
             "CustomerName", "ContactPhones", "Col8", "Col9", "Col10", "Col11",
             "AwaitingPayment", "PublicComments", "POs", "Items"]
        ]
        return out

    def validate_columns(df: pd.DataFrame):
        return [c for c in REQUIRED_COLS if c not in df.columns]

    if not file:
        return

    try:
        raw_df, converted = read_excel_any(file, return_converted_bytes=True)

        if converted:
            st.info("🔁 检测到 HTML/CSV 伪装的 Excel，已自动转换为真实 .xlsx。")
            st.download_button(
                "📥 下载自动转换的 .xlsx",
                converted,
                file_name="converted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        missing = validate_columns(raw_df)
        if missing:
            st.error("❌ 缺少以下必要列，请在原表中补齐后再上传：\n\n- " + "\n- ".join(missing))
            return

        with st.spinner("Processing…"):
            merged = consolidate(raw_df)

        st.success(f"✅ 处理完成，共 {len(merged)} 条订单（每个 OrderNumber 一行）。")
        st.dataframe(merged.head(50), use_container_width=True)

        tsv_no_header = merged.to_csv(sep="\t", header=False, index=False)
        components.html(f"""
            <textarea id="mergedTSV" style="position:absolute;left:-10000px;top:-10000px">{tsv_no_header}</textarea>
            <button onclick="(function(){{
                var t=document.getElementById('mergedTSV');
                t.select();
                document.execCommand('copy');
                alert('✅ 已复制全表（不含表头），可直接粘贴到 Excel/Google Sheets');
            }})()" style="margin:8px 0; padding:.5em 1em; border-radius:8px; border:1px solid #ccc;">
                📋 Copy table (no headers)
            </button>
        """, height=40)

        out_io = BytesIO()
        with pd.ExcelWriter(out_io, engine="xlsxwriter") as writer:
            merged.to_excel(writer, index=False, sheet_name="Consolidated")
        out_io.seek(0)

        st.download_button(
            "📥 Download Merged Excel",
            data=out_io,
            file_name="order_merge_v2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except RuntimeError as e:
        st.error(f"❌ {e}")
    except Exception as e:
        st.error(f"❌ Error: {e}")
