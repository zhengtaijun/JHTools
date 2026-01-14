import re
from io import BytesIO

import pandas as pd
import streamlit as st

from utils.excel_io import read_excel_any
from utils.product_matcher import match_cbm


def render():
    st.subheader("🚚 Advanced Freight Compare + Volume")
    st.markdown("""
本工具对比 **仓库发货表（A）** 与 **自制订单表（B）**：

- 自动识别：A 表含 `First Receipt Date` 列；B 表不含  
- A 表（PO Detail Report）需要列：`PO No`, `Short Description`, `Order Qty`  
- B 表（Fulfilment Report）需要列：`Product_Description`, `SourceFrom`, `qtyRequired`, `OrderNumber`  
- 对比结果分五种情况并合并为一个表：  
  1️⃣ PO + 产品 + 数量完全匹配（仓库 & 我方一致）  
  2️⃣ PO + 产品 匹配，但数量出错  
  3️⃣ 只有 A 有（仓库多做了 / 我方漏单）  
  4️⃣ 只有 B 有（我方下单了 / 仓库漏做）  
  5️⃣ 双方都没有 PO（店内库存 / 展品，无需仓库发货）  
- 使用产品名称匹配 `product_info.xlsx` 中 CBM，计算体积与总和
    """)

    def find_col(df: pd.DataFrame, targets, required=True):
        if isinstance(targets, str):
            targets = [targets]
        cols_lower = [str(c).strip().lower() for c in df.columns]
        for t in targets:
            t_low = t.lower()
            for i, c in enumerate(cols_lower):
                if c == t_low:
                    return i
        for t in targets:
            t_low = t.lower()
            for i, c in enumerate(cols_lower):
                if t_low in c:
                    return i
        if required:
            raise ValueError(f"找不到列：{targets}")
        return None

    RE_FLOAT_INT = re.compile(r"^\s*(\d+)(?:\.0+)?\s*$")
    RE_HASH_PO = re.compile(r"#\s*(\d+)")
    RE_NEED_ORDER = re.compile(r"(on[- ]order|pending)", re.IGNORECASE)

    def normalize_po(value):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return ""
        s = str(value).strip()
        if not s:
            return ""
        m = RE_FLOAT_INT.match(s)
        if m:
            return f"PO{m.group(1)}"
        s_u = s.upper()
        if s_u.startswith("PO"):
            tail = s_u[2:].strip()
            m2 = RE_FLOAT_INT.match(tail)
            if m2:
                return f"PO{m2.group(1)}"
            return s_u
        return "PO" + s

    fileA = st.file_uploader("📄 Upload **PO Detail Report** (with 'First Receipt Date')", type=["xlsx", "xls"], key="freight_A")
    fileB = st.file_uploader("📄 Upload **Fulfilment Report**", type=["xlsx", "xls"], key="freight_B")

    if not (fileA and fileB):
        return

    if st.button("🔍 Compare & Calculate Volume"):
        try:
            dfA = read_excel_any(fileA)
            dfB = read_excel_any(fileB)

            colsA = [str(c) for c in dfA.columns]
            colsB = [str(c) for c in dfB.columns]

            has_first_A = any("first receipt date" in str(c).lower() for c in colsA)
            has_first_B = any("first receipt date" in str(c).lower() for c in colsB)

            if has_first_B and not has_first_A:
                dfA, dfB = dfB, dfA
                colsA, colsB = colsB, colsA
                st.info("ℹ️ 检测到第二个文件才包含 `First Receipt Date`，已自动将其视为表格 A（仓库表）。")
            elif not has_first_A:
                st.warning("⚠️ 未在任一文件中发现 `First Receipt Date` 列，请确认文件是否上传正确。")

            idxA_po = find_col(dfA, ["PO No", "PONo", "PO_Number"])
            idxA_desc = find_col(dfA, ["Short Description", "Short_Description", "Description"])
            idxA_qty = find_col(dfA, ["Order Qty", "OrderQty", "Qty"])

            idxB_desc = find_col(dfB, ["Product_Description", "Product Description", "Product"])
            idxB_source = find_col(dfB, ["SourceFrom", "Source From"])
            idxB_qty = find_col(dfB, ["qtyRequired", "Qty Required", "Order Qty", "OrderQty"])
            idxB_order = find_col(dfB, ["OrderNumber", "Order Number", "OrderNo"])

            with st.expander("👀 Preview A (warehouse)", expanded=False):
                st.write(dfA.head())
            with st.expander("👀 Preview B (internal)", expanded=False):
                st.write(dfB.head())

            rowsA = []
            for _, r in dfA.iterrows():
                po_norm = normalize_po(r.iloc[idxA_po])
                has_po = bool(po_norm)
                desc = str(r.iloc[idxA_desc]) if not pd.isna(r.iloc[idxA_desc]) else ""
                desc_norm = re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", desc.lower())).strip()
                qty_raw = r.iloc[idxA_qty]
                try:
                    qty = int(float(qty_raw)) if str(qty_raw).strip() != "" else 0
                except Exception:
                    qty = 0
                rowsA.append(dict(po=po_norm, has_po=has_po, desc=desc, desc_norm=desc_norm, qty=qty))

            rowsB = []
            for _, r in dfB.iterrows():
                src = "" if pd.isna(r.iloc[idxB_source]) else str(r.iloc[idxB_source])
                need_order = bool(RE_NEED_ORDER.search(src.lower()))
                m_po = RE_HASH_PO.search(src)
                po_norm = ""
                has_po = False
                if need_order and m_po:
                    po_norm = normalize_po(m_po.group(1))
                    has_po = bool(po_norm)

                desc = "" if pd.isna(r.iloc[idxB_desc]) else str(r.iloc[idxB_desc])
                desc_norm = re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", desc.lower())).strip()

                qty_raw = r.iloc[idxB_qty]
                try:
                    qty = int(float(qty_raw)) if str(qty_raw).strip() != "" else 0
                except Exception:
                    qty = 0

                order_no = "" if pd.isna(r.iloc[idxB_order]) else str(r.iloc[idxB_order]).strip()
                rowsB.append(dict(po=po_norm, has_po=has_po, desc=desc, desc_norm=desc_norm, qty=qty, order_no=order_no))

            A_group = {}
            for ra in rowsA:
                if not ra["has_po"]:
                    continue
                key = (ra["po"], ra["desc_norm"])
                A_group.setdefault(key, {"po": ra["po"], "desc": ra["desc"], "qty": 0})
                A_group[key]["qty"] += ra["qty"]

            B_group = {}
            for rb in rowsB:
                if not rb["has_po"]:
                    continue
                key = (rb["po"], rb["desc_norm"])
                B_group.setdefault(key, {"po": rb["po"], "desc": rb["desc"], "qty": 0, "orders": []})
                B_group[key]["qty"] += rb["qty"]
                if rb["order_no"]:
                    B_group[key]["orders"].append(rb["order_no"])

            part1, part2 = [], []
            common_keys = set(A_group.keys()) & set(B_group.keys())
            for key in common_keys:
                ga, gb = A_group[key], B_group[key]
                qty_a, qty_b = ga["qty"], gb["qty"]

                po_cell = ga["po"]
                if gb["orders"]:
                    po_cell = po_cell + ", " + ", ".join(gb["orders"])

                product = gb["desc"] or ga["desc"]

                if qty_a == qty_b:
                    part1.append(dict(Category="1. Match (A & B)", PO_Order=po_cell, Product=product, Qty=qty_a))
                else:
                    part2.append(dict(Category="2. Qty mismatch (PO & product same)", PO_Order=po_cell, Product=product, Qty=qty_a))

            part3 = []
            for key in set(A_group.keys()) - set(B_group.keys()):
                ga = A_group[key]
                part3.append(dict(Category="3. Only in A (warehouse extra)", PO_Order=ga["po"], Product=ga["desc"], Qty=ga["qty"]))

            part4 = []
            for key in set(B_group.keys()) - set(A_group.keys()):
                gb = B_group[key]
                po_cell = gb["po"]
                if gb["orders"]:
                    po_cell = po_cell + ", " + ", ".join(gb["orders"])
                part4.append(dict(Category="4. Only in B (our order, warehouse missing)", PO_Order=po_cell, Product=gb["desc"], Qty=gb["qty"]))

            part5 = []
            for rb in rowsB:
                if rb["has_po"]:
                    continue
                part5.append(dict(Category="5. No PO (store stock / display)", PO_Order=rb["order_no"], Product=rb["desc"], Qty=rb["qty"]))

            all_rows = part1 + part2 + part3 + part4 + part5
            if not all_rows:
                st.error("未找到任何记录，请检查两个表格内容是否正确。")
                return

            df_result = pd.DataFrame(all_rows)
            df_result["Volume"] = [match_cbm(x or "") for x in df_result["Product"].tolist()]
            df_result["Qty"] = pd.to_numeric(df_result["Qty"], errors="coerce").fillna(0)
            df_result["Total Volume"] = df_result["Volume"] * df_result["Qty"]

            total_volume_sum = float(df_result["Total Volume"].sum())

            summary_row = {"Category": "TOTAL", "PO_Order": "", "Product": "", "Qty": "", "Volume": "", "Total Volume": total_volume_sum}
            df_final = pd.concat([df_result, pd.DataFrame([summary_row])], ignore_index=True)

            st.success(f"✅ Completed. Total rows: {len(df_result)}, Total Volume: **{total_volume_sum:.3f} m³**")

            def show(cat, title):
                st.markdown(title)
                st.dataframe(
                    df_result[df_result["Category"] == cat][["PO_Order", "Product", "Qty", "Volume", "Total Volume"]],
                    use_container_width=True,
                )

            show("1. Match (A & B)", "### 📊 Part 1 – 完全匹配（仓库 & 我方一致）")
            show("2. Qty mismatch (PO & product same)", "### 📊 Part 2 – 数量不一致（PO & 产品相同）")
            show("3. Only in A (warehouse extra)", "### 📊 Part 3 – 仅 A 有（仓库多做 / 我方漏单）")
            show("4. Only in B (our order, warehouse missing)", "### 📊 Part 4 – 仅 B 有 PO（我方下单 / 仓库漏做）")
            show("5. No PO (store stock / display)", "### 📊 Part 5 – 无 PO（店内库存 / 展品）")

            out = BytesIO()
            with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
                df_final.to_excel(writer, index=False, sheet_name="Compare")
                workbook = writer.book
                worksheet = writer.sheets["Compare"]

                fmt_part1 = workbook.add_format({"bg_color": "#C6EFCE"})
                fmt_part2 = workbook.add_format({"bg_color": "#F8CBAD"})
                fmt_part3 = workbook.add_format({"bg_color": "#FFEB9C"})
                fmt_part4 = workbook.add_format({"bg_color": "#FFC7CE"})
                fmt_part5 = workbook.add_format({"bg_color": "#D9E1F2"})
                fmt_total = workbook.add_format({"bold": True})

                cat_col_idx = df_final.columns.get_loc("Category")
                for row_idx in range(1, len(df_final) + 1):
                    cat = df_final.iloc[row_idx - 1, cat_col_idx]
                    fmt = None
                    if str(cat).startswith("1. "):
                        fmt = fmt_part1
                    elif str(cat).startswith("2. Qty"):
                        fmt = fmt_part2
                    elif str(cat).startswith("3. Only in A"):
                        fmt = fmt_part3
                    elif str(cat).startswith("4. Only in B"):
                        fmt = fmt_part4
                    elif str(cat).startswith("5. No PO"):
                        fmt = fmt_part5
                    elif cat == "TOTAL":
                        fmt = fmt_total
                    if fmt:
                        worksheet.set_row(row_idx, None, fmt)

            out.seek(0)
            st.download_button(
                "📥 Download Excel (with colored sections)",
                out,
                file_name="freight_compare_volume.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"❌ Error: {e}")
