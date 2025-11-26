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
import re
from functools import lru_cache
from rapidfuzz import process, fuzz



# -------- Robust Excel/HTML/CSV loader --------
# -------- Robust Excel/HTML/CSV loader with auto-convert-to-xlsx --------

def _ensure_xlrd_ok():
    try:
        import xlrd
        parts = tuple(int(p) for p in xlrd.__version__.split(".")[:3])
        if parts < (2, 0, 1):
            raise RuntimeError(f"xlrd {xlrd.__version__} too old; please upgrade to xlrd>=2.0.1")
    except ImportError:
        raise RuntimeError("xlrd not installed; please `pip install xlrd>=2.0.1`")

def _to_xlsx_bytes(df: pd.DataFrame) -> BytesIO:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    bio.seek(0)
    return bio

def read_excel_any(file_obj, return_converted_bytes: bool = False, **kwargs):
    name = (getattr(file_obj, "name", "") or "").lower()

    raw = file_obj.read() if hasattr(file_obj, "read") else file_obj
    if not isinstance(raw, (bytes, bytearray)):
        try:
            file_obj.seek(0)
            raw = file_obj.read()
        except Exception:
            df = pd.read_excel(file_obj, **kwargs)
            return (df, None) if return_converted_bytes else df

    data = bytes(raw)
    head = data[:64]
    sniff = data[:2048].lower()  # 多 sniff 一点，便于发现 <table> 在前

    def as_bio():
        return BytesIO(data)

    # ---------- 1) HTML 伪装的 Excel：<html / <!doctype / <table 都算 ----------
    if (sniff.lstrip().startswith(b"<html")
        or sniff.lstrip().startswith(b"<!doctype html")
        or b"<table" in sniff):          # 关键：补上这条
        # 用 read_html 解析，不把第一行当表头
        tables = pd.read_html(as_bio(), header=None)
        if not tables:
            raise RuntimeError("HTML 文件中未发现可解析的表格。请导出为真正的 Excel。")
        df = tables[0]

        # 如果第一行包含你的标准字段，把第一行提为列名
        expected_cols = {
            "datecreated","ordernumber","orderstatus","product_description","size",
            "colour","customername","phone","mobile","deliverymode",
            "publiccomments","qtyrequired","sourcefrom"
        }
        first_row = [str(x).strip() for x in df.iloc[0].tolist()]
        if any(x.lower() in expected_cols for x in first_row):
            df.columns = df.iloc[0]
            df = df.drop(df.index[0]).reset_index(drop=True)

        # 统一成字符串（和 dtype=str 效果一致）
        df = df.applymap(lambda x: "" if pd.isna(x) else str(x))

        conv = _to_xlsx_bytes(df) if return_converted_bytes else None
        return (df, conv) if return_converted_bytes else df

    # ---------- 2) 真 .xlsx（ZIP 头） ----------
    if head.startswith(b"PK\x03\x04"):
        try:
            df = pd.read_excel(as_bio(), engine="openpyxl", **kwargs)
        except Exception:
            df = pd.read_excel(as_bio(), **kwargs)
        return (df, None) if return_converted_bytes else df

    # ---------- 3) 真 .xls（OLE2 头）或扩展名 .xls ----------
    if head.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1") or name.endswith(".xls"):
        _ensure_xlrd_ok()
        df = pd.read_excel(as_bio(), engine="xlrd", **kwargs)
        return (df, None) if return_converted_bytes else df

    # ---------- 4) CSV/TSV 误扩展 ----------
    text_sample = data[:4096].decode("utf-8", errors="ignore")
    if ("\t" in text_sample or "," in text_sample) and ("\n" in text_sample or "\r" in text_sample):
        sep = "\t" if text_sample.count("\t") >= text_sample.count(",") else ","
        df = pd.read_csv(BytesIO(data), sep=sep)
        df = df.applymap(lambda x: "" if pd.isna(x) else str(x))
        conv = _to_xlsx_bytes(df) if return_converted_bytes else None
        return (df, conv) if return_converted_bytes else df

    # ---------- 5) 兜底 ----------
    df = pd.read_excel(as_bio(), **kwargs)
    return (df, None) if return_converted_bytes else df


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
    ["TRF Volume Calculator", "Order Merge Tool", "Order Merge Tool V2", "Profit Calculator", "List Split", "Image Table Extractor", "Google Sheet Query"],
    index=0
)

# ========== TOOL 1: TRF Volume Calculator ==========
if tool == "TRF Volume Calculator":
    st.subheader("📦 TRF Volume Calculator")
    st.markdown("📺 [Instructional video](https://youtu.be/S10a3kPEXZg)")

    PRODUCT_INFO_URL = (
        "https://raw.githubusercontent.com/zhengtaijun/JHCH_TRF-Volume/main/product_info.xlsx"
    )

    # ===================== 统一标准化与别名归一（保持原有逻辑） =====================

    _WS_RE = re.compile(r"\s+")
    _PUNCT_RE = re.compile(r"[^a-z0-9]+")

    ALIASES = {
        "drawer": ["drawers", "drw", "drws"],
        "tallboy": ["tall boy", "tall-boy"],
        "queen": ["qn", "qs", "queen-size", "queen size"],
        "king": ["kg", "ks", "king-size", "king size"],
    }

    def _apply_aliases(tokens):
        out = []
        for t in tokens:
            replaced = False
            for canon, variants in ALIASES.items():
                if t == canon or t in variants:
                    out.append(canon)
                    replaced = True
                    break
            if not replaced:
                out.append(t)
        return out

    def normalize(s: str) -> str:
        s = s.strip().lower()
        s = _PUNCT_RE.sub(" ", s)
        s = _WS_RE.sub(" ", s)
        tokens = s.split()
        tokens = _apply_aliases(tokens)
        return " ".join(tokens)

    def fingerprint(s: str) -> str:
        toks = normalize(s).split()
        return " ".join(sorted(set(toks)))

    @st.cache_data
    def load_product_info_and_build_index():
        resp = requests.get(PRODUCT_INFO_URL)
        resp.raise_for_status()
        df = read_excel_any(BytesIO(resp.content))

        with st.expander("✅ Product-info file loaded. Click to view columns", expanded=False):
            st.write(df.columns.tolist())

        if {"Product Name", "CBM"} - set(df.columns):
            raise ValueError("`Product Name` and `CBM` columns are required.")

        names = df["Product Name"].fillna("").astype(str).tolist()
        cbms = pd.to_numeric(df["CBM"], errors="coerce").fillna(0).tolist()

        product_dict_raw = dict(zip(names, cbms))

        norm_index = {}
        fp_index = {}
        names_norm_list = []

        for n, c in zip(names, cbms):
            n_norm = normalize(n)
            n_fp = " ".join(sorted(set(n_norm.split())))
            norm_index[n_norm] = c
            fp_index[n_fp] = c
            names_norm_list.append(n_norm)

        return {
            "df": df,
            "product_dict_raw": product_dict_raw,
            "norm_index": norm_index,
            "fp_index": fp_index,
            "names_norm_list": names_norm_list,
            "names_all": names,
            "cbms_all": cbms,
        }

    idx = load_product_info_and_build_index()

    @lru_cache(maxsize=4096)
    def match_product(name: str):
        if not name:
            return None

        raw = idx["product_dict_raw"].get(name)
        if raw is not None:
            return raw

        n_norm = normalize(name)
        got = idx["norm_index"].get(n_norm)
        if got is not None:
            return got

        n_fp = " ".join(sorted(set(n_norm.split())))
        got = idx["fp_index"].get(n_fp)
        if got is not None:
            return got

        tokens = n_norm.split()
        prefix = " ".join(tokens[:3]) if len(tokens) >= 3 else " ".join(tokens)
        if prefix:
            m_prefix = process.extractOne(
                prefix,
                [" ".join(t.split()[:3]) for t in idx["names_norm_list"]],
                scorer=fuzz.token_set_ratio,
                score_cutoff=90,
            )
            if m_prefix:
                _, _, matched_idx = m_prefix
                return idx["cbms_all"][matched_idx]

        m1 = process.extractOne(
            n_norm, idx["names_norm_list"], scorer=fuzz.token_set_ratio, score_cutoff=88
        )
        if m1:
            _, _, matched_idx = m1
            return idx["cbms_all"][matched_idx]

        m2 = process.extractOne(
            n_norm, idx["names_norm_list"], scorer=fuzz.partial_ratio, score_cutoff=85
        )
        if m2:
            _, _, matched_idx = m2
            return idx["cbms_all"][matched_idx]

        return None

    # ===================== 上传文件 + 自动识别列 =====================
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

    # ---- 用户可编辑的表头关键字（默认用你的标题） ----
    st.markdown("### 🔠 Column headers (auto-detect, editable)")
    header_prod = st.text_input("Header for **Product Name**", value="Short Description")
    header_order = st.text_input("Header for **Order Number** (Invoices)", value="Invoices")
    header_qty = st.text_input("Header for **Quantity**", value="Order Qty")
    header_po = st.text_input("Header for **PO Number**", value="PO No")

    # ---- 根据表头尝试自动检测列号，用于给 number_input 预填值 ----
    def _auto_detect_col(df, header_name, fallback=1):
        if df is None or not header_name:
            return fallback
        cols_lower = [str(c).strip().lower() for c in df.columns]
        target = header_name.strip().lower()
        # 先完全匹配
        for i, c in enumerate(cols_lower):
            if c == target:
                return i + 1  # 1-based
        # 再做包含匹配
        for i, c in enumerate(cols_lower):
            if target in c:
                return i + 1
        return fallback

    if n_cols == 0:
        max_cols = 50
    else:
        max_cols = n_cols

    auto_prod = _auto_detect_col(df_preview, header_prod, fallback=3)
    auto_order = _auto_detect_col(df_preview, header_order, fallback=7)
    auto_qty = _auto_detect_col(df_preview, header_qty, fallback=8)
    auto_po = _auto_detect_col(df_preview, header_po, fallback=1)

    # ---- 手动列号（1-based），默认显示自动匹配到的值，仍可修改 ----
    # ---- 手动列号（1-based），默认显示自动匹配到的值，仍可修改 ----
    st.markdown("### #️⃣ Column numbers (1-based, optional override)")

    # 开关：是否用手动列号强行 override
    use_manual_cols = st.checkbox(
        "Use manual column numbers to override header detection",
        value=False,
        help="默认关闭：仅使用上面的表头标题来识别列。勾选后：强制使用下面的列号。"
    )

    col_prod = st.number_input(
        "Column # of **Product Name**",
        min_value=1,
        max_value=max_cols,
        value=min(max(auto_prod, 1), max_cols),
    )
    col_order = st.number_input(
        "Column # of **Order Number (Invoices)**",
        min_value=1,
        max_value=max_cols,
        value=min(max(auto_order, 1), max_cols),
    )
    col_qty = st.number_input(
        "Column # of **Quantity**",
        min_value=1,
        max_value=max_cols,
        value=min(max(auto_qty, 1), max_cols),
    )
    col_po = st.number_input(
        "Column # of **PO Number**",
        min_value=1,
        max_value=max_cols,
        value=min(max(auto_po, 1), max_cols),
    )

    # ---- 计算时实际决定使用哪一列 ----
    def resolve_col_index(df, header_name, manual_1based, auto_1based, use_manual: bool):
        """
        返回 0-based 列索引：

        - 当 use_manual=True 时：
            直接使用 manual_1based - 1（强制 override）
        - 当 use_manual=False 时：
            只用 header_name 在表头中查找（先全等后包含），
            如果完全找不到，则退回自动侦测的 auto_1based - 1
        """
        # 手动 override 模式
        if use_manual and manual_1based is not None:
            return int(manual_1based) - 1

        # 标题匹配模式（默认）
        if df is not None and header_name:
            cols_lower = [str(c).strip().lower() for c in df.columns]
            target = header_name.strip().lower()
            # 完全匹配
            for i, c in enumerate(cols_lower):
                if c == target:
                    return i
            # 包含匹配
            for i, c in enumerate(cols_lower):
                if target and target in c:
                    return i

        # 标题完全找不到，退回自动识别出来的列号
        if auto_1based is not None:
            return int(auto_1based) - 1

        raise ValueError(f"Cannot resolve column for header '{header_name}'")


    # ===================== 体积计算流程（带 PO + Invoices + 精简列） =====================
    # ===================== 体积计算流程（带 PO + Short Description + 精简列） =====================
    def process_volume_file(file_bytes, prod_idx, qty_idx, inv_idx, po_idx):
        dfw = read_excel_any(BytesIO(file_bytes))

        # 安全保护：索引溢出就报错
        ncols = dfw.shape[1]
        for idx_check, label in [
            (prod_idx, "Product Name"),
            (qty_idx, "Quantity"),
            (inv_idx, "Invoices"),
            (po_idx, "PO No"),
        ]:
            if idx_check < 0 or idx_check >= ncols:
                raise ValueError(f"Column index for {label} out of range.")

        # 基础列
        product_series = dfw.iloc[:, prod_idx].fillna("").astype(str)
        product_names = product_series.tolist()
        quantities = pd.to_numeric(dfw.iloc[:, qty_idx], errors="coerce").fillna(0)

        inv_series = dfw.iloc[:, inv_idx] if inv_idx is not None else pd.Series([""] * len(dfw))
        po_series = dfw.iloc[:, po_idx] if po_idx is not None else pd.Series([""] * len(dfw))

        # 合并 PO No 与 Invoices 到一个单元格（PO 在前，用逗号隔开，并在数字前加 "PO"）
        merged_ref = []
        for po, inv in zip(po_series, inv_series):
            # --- 处理 PO 为纯文本，去掉 .0、小数、空格等 ---
            po_s = "" if pd.isna(po) else str(po).strip()

            if po_s:
                # 如果是纯数字或数字.0，变成纯整数字符串
                if re.fullmatch(r"\d+(\.0+)?", po_s):
                    po_s = po_s.split(".", 1)[0]

                po_s = po_s.strip()

                # 统一加上“PO”前缀（如果没有）
                if not po_s.upper().startswith("PO"):
                    po_s = f"PO{po_s}"

            inv_s = "" if pd.isna(inv) else str(inv).strip()

            if po_s and inv_s:
                merged_ref.append(f"{po_s}, {inv_s}")
            elif po_s:
                merged_ref.append(po_s)
            else:
                merged_ref.append(inv_s)


        # ---- 并行做体积匹配 ----
        total = len(product_names)
        volumes = [None] * total

        def worker(start: int, end: int):
            out = []
            for i in range(start, end):
                nm = product_names[i].strip()
                vol = match_product(nm) if nm else None
                out.append(vol)
            return out

        with ThreadPoolExecutor(max_workers=4) as pool:
            chunk = max(total // 4, 1)
            futures = [
                pool.submit(worker, i * chunk, (i + 1) * chunk if i < 3 else total)
                for i in range(4)
            ]
            pos = 0
            for f in futures:
                batch = f.result()
                volumes[pos : pos + len(batch)] = batch
                pos += len(batch)

        # 只保留三列 + Volume & Total Volume
        df_res = pd.DataFrame(
            {
                "PO/Invoice": merged_ref,                     # 合并 PO No + Invoices
                "Short Description": product_series,          # 产品名列
                "Order Qty": quantities,                      # 数量
            }
        )

        df_res["Volume"] = pd.to_numeric(pd.Series(volumes), errors="coerce").fillna(0)
        df_res["Total Volume"] = df_res["Volume"] * df_res["Order Qty"]

        # 最后一行汇总 Total Volume
        summary = pd.DataFrame(
            {
                "PO/Invoice": [""],
                "Short Description": [""],
                "Order Qty": [""],
                "Volume": [""],
                "Total Volume": [df_res["Total Volume"].sum()],
            }
        )

        df_final = pd.concat([df_res, summary], ignore_index=True)
        return df_final


    # ===================== 触发计算与下载 =====================
    if warehouse_file and uploaded_bytes and st.button("Calculate volume"):
        if df_preview is None:
            st.error("❌ File not loaded correctly, please re-upload.")
        else:
            with st.spinner("Processing…"):
                try:
                    prod_idx = resolve_col_index(df_preview, header_prod, col_prod, auto_prod, use_manual_cols)
                    qty_idx  = resolve_col_index(df_preview, header_qty,  col_qty,  auto_qty,  use_manual_cols)
                    inv_idx  = resolve_col_index(df_preview, header_order, col_order, auto_order, use_manual_cols)
                    po_idx   = resolve_col_index(df_preview, header_po,   col_po,   auto_po,   use_manual_cols)


                    result_df = process_volume_file(
                        uploaded_bytes, prod_idx, qty_idx, inv_idx, po_idx
                    )

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

    pass

# ========== TOOL 2: Order Merge Tool ==========
# ========== TOOL 2: Order Freight Compare & Volume (New) ==========
elif tool == "Order Merge Tool":
    st.subheader("🚚 Advanced Freight Compare + Volume")
    st.markdown("""
    本工具对比 **仓库发货表（A）** 与 **自制订单表（B）**：

    - 自动识别：A 表含 `First Receipt Date` 列；B 表不含  
    - A 表需要列：`PO No`, `Short Description`, `Order Qty`  
    - B 表需要列：`Product_Description`, `SourceFrom`, `qtyRequired`, `OrderNumber`  
    - 对比结果分四种情况并合并为一个表：  
      1️⃣ PO + 产品 + 数量完全匹配（仓库 & 我方一致）  
      2️⃣ 只有 A 有（仓库多做了 / 我方漏单）  
      3️⃣ 只有 B 有（我方下单了 / 仓库漏做）  
      4️⃣ 双方都没有 PO（店内库存 / 展品，无需仓库发货）  
    - 使用产品名称匹配 `product_info.xlsx` 中 CBM，计算体积与总和
    """)

    # ---------- 共用：product_info 体积字典 & 模糊匹配 ----------
    PRODUCT_INFO_URL = (
        "https://raw.githubusercontent.com/zhengtaijun/JHCH_TRF-Volume/main/product_info.xlsx"
    )

    _WS_RE = re.compile(r"\s+")
    _PUNCT_RE = re.compile(r"[^a-z0-9]+")

    ALIASES = {
        "drawer": ["drawers", "drw", "drws"],
        "tallboy": ["tall boy", "tall-boy"],
        "queen": ["qn", "qs", "queen-size", "queen size"],
        "king": ["kg", "ks", "king-size", "king size"],
    }

    def _apply_aliases(tokens):
        out = []
        for t in tokens:
            replaced = False
            for canon, variants in ALIASES.items():
                if t == canon or t in variants:
                    out.append(canon)
                    replaced = True
                    break
            if not replaced:
                out.append(t)
        return out

    def normalize_name(s: str) -> str:
        s = (str(s) if s is not None else "").strip().lower()
        s = _PUNCT_RE.sub(" ", s)
        s = _WS_RE.sub(" ", s)
        tokens = s.split()
        tokens = _apply_aliases(tokens)
        return " ".join(tokens)

    @st.cache_data
    def load_product_info_index():
        resp = requests.get(PRODUCT_INFO_URL)
        resp.raise_for_status()
        df = read_excel_any(BytesIO(resp.content))

        if {"Product Name", "CBM"} - set(df.columns):
            raise ValueError("`Product Name` 和 `CBM` 列在 product_info.xlsx 中是必须的。")

        names = df["Product Name"].fillna("").astype(str).tolist()
        cbms = pd.to_numeric(df["CBM"], errors="coerce").fillna(0).tolist()

        product_dict_raw = dict(zip(names, cbms))

        norm_index = {}
        fp_index = {}
        names_norm_list = []

        for n, c in zip(names, cbms):
            n_norm = normalize_name(n)
            n_fp = " ".join(sorted(set(n_norm.split())))
            norm_index[n_norm] = c
            fp_index[n_fp] = c
            names_norm_list.append(n_norm)

        return {
            "df": df,
            "product_dict_raw": product_dict_raw,
            "norm_index": norm_index,
            "fp_index": fp_index,
            "names_norm_list": names_norm_list,
            "names_all": names,
            "cbms_all": cbms,
        }

    idx_vol = load_product_info_index()

    @lru_cache(maxsize=4096)
    def match_product_cmb(name: str):
        """根据产品名匹配 CBM（复制自 TRF Volume 逻辑）"""
        if not name:
            return None

        # 0. 原文精确
        raw = idx_vol["product_dict_raw"].get(name)
        if raw is not None:
            return raw

        n_norm = normalize_name(name)

        # 1. 规范化精确
        got = idx_vol["norm_index"].get(n_norm)
        if got is not None:
            return got

        # 2. 指纹精确
        n_fp = " ".join(sorted(set(n_norm.split())))
        got = idx_vol["fp_index"].get(n_fp)
        if got is not None:
            return got

        # 3a. 前缀模糊
        tokens = n_norm.split()
        prefix = " ".join(tokens[:3]) if len(tokens) >= 3 else " ".join(tokens)
        if prefix:
            m_prefix = process.extractOne(
                prefix,
                [" ".join(t.split()[:3]) for t in idx_vol["names_norm_list"]],
                scorer=fuzz.token_set_ratio,
                score_cutoff=90,
            )
            if m_prefix:
                _, _, matched_idx = m_prefix
                return idx_vol["cbms_all"][matched_idx]

        # 3b. 全名模糊
        m1 = process.extractOne(
            n_norm, idx_vol["names_norm_list"], scorer=fuzz.token_set_ratio, score_cutoff=88
        )
        if m1:
            _, _, matched_idx = m1
            return idx_vol["cbms_all"][matched_idx]

        # 3c. partial 兜底
        m2 = process.extractOne(
            n_norm, idx_vol["names_norm_list"], scorer=fuzz.partial_ratio, score_cutoff=85
        )
        if m2:
            _, _, matched_idx = m2
            return idx_vol["cbms_all"][matched_idx]

        return None

    # ---------- 工具函数：列名自动匹配 ----------
    def find_col(df: pd.DataFrame, targets, required=True):
        """
        在 df.columns 中查找列名，targets 可以是字符串或列表（不区分大小写）；
        先全等，再包含；找不到且 required=True 则报错。
        返回 0-based index。
        """
        if isinstance(targets, str):
            targets = [targets]
        cols_lower = [str(c).strip().lower() for c in df.columns]
        for t in targets:
            t_low = t.lower()
            # 完全匹配
            for i, c in enumerate(cols_lower):
                if c == t_low:
                    return i
        for t in targets:
            t_low = t.lower()
            # 包含匹配
            for i, c in enumerate(cols_lower):
                if t_low in c:
                    return i
        if required:
            raise ValueError(f"找不到列：{targets}")
        return None

    # ---------- 工具函数：规范 PO 编号为文本 "POxxxx" ----------
    RE_FLOAT_INT = re.compile(r"^\s*(\d+)(?:\.0+)?\s*$")
    RE_HASH_PO = re.compile(r"#\s*(\d+)")
    RE_NEED_ORDER = re.compile(r"(on[- ]order|pending)", re.IGNORECASE)

    def normalize_po(value):
        """把各种格式（1234 / 1234.0 / 'PO1234'）统一为 'PO1234' 文本；空返回 ''。"""
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return ""
        s = str(value).strip()
        if not s:
            return ""
        # 纯数字或类似 1234.0
        m = RE_FLOAT_INT.match(s)
        if m:
            num = m.group(1)
            return f"PO{num}"
        # 已有 PO 前缀
        s_u = s.upper()
        if s_u.startswith("PO"):
            # 去掉可能的 PO00123.0 这类
            tail = s_u[2:].strip()
            m2 = RE_FLOAT_INT.match(tail)
            if m2:
                return f"PO{m2.group(1)}"
            return s_u
        # 其他情况当普通文本，加 PO
        return "PO" + s

    # ---------- 上传两个文件 ----------
    fileA = st.file_uploader("📄 Upload **Warehouse file A** (with 'First Receipt Date')", type=["xlsx", "xls"], key="freight_A")
    fileB = st.file_uploader("📄 Upload **Internal order file B**", type=["xlsx", "xls"], key="freight_B")

    if fileA and fileB and st.button("🔍 Compare & Calculate Volume"):
        try:
            # 读取
            dfA = read_excel_any(fileA)
            dfB = read_excel_any(fileB)

            colsA = [str(c) for c in dfA.columns]
            colsB = [str(c) for c in dfB.columns]

            has_first_A = any("first receipt date" in str(c).lower() for c in colsA)
            has_first_B = any("first receipt date" in str(c).lower() for c in colsB)

            # 自动纠正 A/B 角色：谁含 First Receipt Date 谁就是 A
            if has_first_B and not has_first_A:
                dfA, dfB = dfB, dfA
                colsA, colsB = colsB, colsA
                st.info("ℹ️ 检测到第二个文件才包含 `First Receipt Date`，已自动将其视为表格 A（仓库表）。")
            elif not has_first_A:
                st.warning("⚠️ 未在任一文件中发现 `First Receipt Date` 列，请确认文件是否上传正确。")

            # ---- 找 A 表列：PO No / Short Description / Order Qty ----
            idxA_po = find_col(dfA, ["PO No", "PONo", "PO_Number"])
            idxA_desc = find_col(dfA, ["Short Description", "Short_Description", "Description"])
            idxA_qty = find_col(dfA, ["Order Qty", "OrderQty", "Qty"])

            # ---- 找 B 表列：Product_Description / SourceFrom / qtyRequired / OrderNumber ----
            idxB_desc = find_col(dfB, ["Product_Description", "Product Description", "Product"])
            idxB_source = find_col(dfB, ["SourceFrom", "Source From"])
            idxB_qty = find_col(dfB, ["qtyRequired", "Qty Required", "Order Qty", "OrderQty"])
            idxB_order = find_col(dfB, ["OrderNumber", "Order Number", "OrderNo"])

            # 预览
            with st.expander("👀 Preview A (warehouse)", expanded=False):
                st.write(dfA.head())
            with st.expander("👀 Preview B (internal)", expanded=False):
                st.write(dfB.head())

            # ---------- 预处理 A 表 ----------
            rowsA = []
            for i, r in dfA.iterrows():
                po_raw = r.iloc[idxA_po]
                po_norm = normalize_po(po_raw)
                has_po = bool(po_norm)

                desc = str(r.iloc[idxA_desc]) if not pd.isna(r.iloc[idxA_desc]) else ""
                desc_norm = normalize_name(desc)

                qty_raw = r.iloc[idxA_qty]
                try:
                    qty = int(float(qty_raw)) if str(qty_raw).strip() != "" else 0
                except Exception:
                    qty = 0

                rowsA.append(
                    dict(
                        idx=i,
                        po=po_norm,
                        has_po=has_po,
                        desc=desc,
                        desc_norm=desc_norm,
                        qty=qty,
                    )
                )

            # ---------- 预处理 B 表 ----------
            rowsB = []
            for i, r in dfB.iterrows():
                src = "" if pd.isna(r.iloc[idxB_source]) else str(r.iloc[idxB_source])
                src_low = src.lower()

                ## 判断是否需要订货：SourceFrom 中包含 On-Order 或 Pending
                need_order = bool(RE_NEED_ORDER.search(src_low))
                m_po = RE_HASH_PO.search(src)
                po_norm = ""
                has_po = False
                if need_order and m_po:
                    po_norm = normalize_po(m_po.group(1))
                    has_po = bool(po_norm)

                desc = "" if pd.isna(r.iloc[idxB_desc]) else str(r.iloc[idxB_desc])
                desc_norm = normalize_name(desc)

                qty_raw = r.iloc[idxB_qty]
                try:
                    qty = int(float(qty_raw)) if str(qty_raw).strip() != "" else 0
                except Exception:
                    qty = 0

                order_no = "" if pd.isna(r.iloc[idxB_order]) else str(r.iloc[idxB_order]).strip()

                rowsB.append(
                    dict(
                        idx=i,
                        po=po_norm,
                        has_po=has_po,
                        desc=desc,
                        desc_norm=desc_norm,
                        qty=qty,
                        order_no=order_no,
                    )
                )

            # ---------- 按 (PO, desc_norm, qty) 建立 key ----------
            # ---------- 先按 (PO, desc_norm) 聚合，方便对比数量 ----------
            # A：只看有 PO 的行
            A_group = {}
            for ra in rowsA:
                if not ra["has_po"]:
                    continue
                key = (ra["po"], ra["desc_norm"])
                if key not in A_group:
                    A_group[key] = {
                        "po": ra["po"],
                        "desc": ra["desc"],
                        "qty": 0,
                    }
                A_group[key]["qty"] += ra["qty"]

            # B：只看有 PO 的行，同时聚合 OrderNumber（多个就合并）
            B_group = {}
            for rb in rowsB:
                if not rb["has_po"]:
                    continue
                key = (rb["po"], rb["desc_norm"])
                if key not in B_group:
                    B_group[key] = {
                        "po": rb["po"],
                        "desc": rb["desc"],
                        "qty": 0,
                        "orders": [],
                    }
                B_group[key]["qty"] += rb["qty"]
                if rb["order_no"]:
                    B_group[key]["orders"].append(rb["order_no"])

            # ---------- Part 1 & Part 2：交集里的 “完全匹配” 和 “数量不一致” ----------
            part1 = []  # 完全匹配
            part2 = []  # 数量不一致（PO + 产品一致，qty 不同）

            common_keys = set(A_group.keys()) & set(B_group.keys())
            for key in common_keys:
                ga = A_group[key]
                gb = B_group[key]
                qty_a = ga["qty"]
                qty_b = gb["qty"]

                # 合并 PO + OrderNumber
                po_cell = ga["po"]
                if gb["orders"]:
                    # 多个订单号用逗号拼在后面
                    po_cell = po_cell + ", " + ", ".join(gb["orders"])

                product = gb["desc"] or ga["desc"]  # 优先用 B 的描述

                if qty_a == qty_b:
                    part1.append(
                        dict(
                            Category="1. Match (A & B)",
                            PO_Order=po_cell,
                            Product=product,
                            Qty=qty_a,
                        )
                    )
                else:
                    # 数量不一致：仍然以 A 的数量作为体积计算的基准
                    part2.append(
                        dict(
                            Category="2. Qty mismatch (PO & product same)",
                            PO_Order=po_cell,
                            Product=product,
                            Qty=qty_a,
                            # 如果后面你想看 B 的数量，也可以额外加一列 Qty_B
                            # Qty_B=qty_b,
                        )
                    )

            # ---------- Part 3：只在 A 有（仓库有，我们没下单 / 下少了） ----------
            part3 = []
            onlyA_keys = set(A_group.keys()) - set(B_group.keys())
            for key in onlyA_keys:
                ga = A_group[key]
                part3.append(
                    dict(
                        Category="3. Only in A (warehouse extra)",
                        PO_Order=ga["po"],
                        Product=ga["desc"],
                        Qty=ga["qty"],
                    )
                )

            # ---------- Part 4：只在 B 有 PO（我们下单，仓库没做） ----------
            part4 = []
            onlyB_keys = set(B_group.keys()) - set(A_group.keys())
            for key in onlyB_keys:
                gb = B_group[key]
                po_cell = gb["po"]
                if gb["orders"]:
                    po_cell = po_cell + ", " + ", ".join(gb["orders"])
                part4.append(
                    dict(
                        Category="4. Only in B (our order, warehouse missing)",
                        PO_Order=po_cell,
                        Product=gb["desc"],
                        Qty=gb["qty"],
                    )
                )

            # ---------- Part 5：双方都没 PO（店内库存 / 展品） ----------
            # 这里仍按 B 表中 “没有 PO 的行” 视为店内库存
            part5 = []
            for rb in rowsB:
                if rb["has_po"]:
                    continue
                part5.append(
                    dict(
                        Category="5. No PO (store stock / display)",
                        PO_Order=rb["order_no"],
                        Product=rb["desc"],
                        Qty=rb["qty"],
                    )
                )

            # ---------- 合并五部分，计算 Volume ----------
            all_rows = part1 + part2 + part3 + part4 + part5
            if not all_rows:
                st.error("未找到任何记录，请检查两个表格内容是否正确。")
            else:
                df_result = pd.DataFrame(all_rows)

                # 体积匹配
                vols = []
                for name in df_result["Product"].tolist():
                    v = match_product_cmb(name or "")
                    vols.append(v if v is not None else 0)
                df_result["Volume"] = pd.to_numeric(pd.Series(vols), errors="coerce").fillna(0)

                df_result["Qty"] = pd.to_numeric(df_result["Qty"], errors="coerce").fillna(0)
                df_result["Total Volume"] = df_result["Volume"] * df_result["Qty"]

                total_volume_sum = df_result["Total Volume"].sum()

                # 最后一行汇总
                summary_row = {
                    "Category": "TOTAL",
                    "PO_Order": "",
                    "Product": "",
                    "Qty": "",
                    "Volume": "",
                    "Total Volume": total_volume_sum,
                }
                df_final = pd.concat([df_result, pd.DataFrame([summary_row])], ignore_index=True)

                st.success(f"✅ Completed. Total rows: {len(df_result)}, Total Volume: **{total_volume_sum:.3f} m³**")

                # 分部分展示
                st.markdown("### 📊 Part 1 – 完全匹配（仓库 & 我方一致）")
                st.dataframe(
                    df_result[df_result["Category"] == "1. Match (A & B)"][["PO_Order","Product","Qty","Volume","Total Volume"]],
                    use_container_width=True,
                )

                st.markdown("### 📊 Part 2 – 数量不一致（PO & 产品相同）")
                st.dataframe(
                    df_result[df_result["Category"] == "2. Qty mismatch (PO & product same)"][["PO_Order","Product","Qty","Volume","Total Volume"]],
                    use_container_width=True,
                )

                st.markdown("### 📊 Part 3 – 仅 A 有（仓库多做 / 我方漏单）")
                st.dataframe(
                    df_result[df_result["Category"] == "3. Only in A (warehouse extra)"][["PO_Order","Product","Qty","Volume","Total Volume"]],
                    use_container_width=True,
                )

                st.markdown("### 📊 Part 4 – 仅 B 有 PO（我方下单 / 仓库漏做）")
                st.dataframe(
                    df_result[df_result["Category"] == "4. Only in B (our order, warehouse missing)"][["PO_Order","Product","Qty","Volume","Total Volume"]],
                    use_container_width=True,
                )

                st.markdown("### 📊 Part 5 – 无 PO（店内库存 / 展品）")
                st.dataframe(
                    df_result[df_result["Category"] == "5. No PO (store stock / display)"][["PO_Order","Product","Qty","Volume","Total Volume"]],
                    use_container_width=True,
                )


                # 导出到带颜色的 Excel
                out = BytesIO()
                with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
                    df_final.to_excel(writer, index=False, sheet_name="Compare")
                    workbook = writer.book
                    worksheet = writer.sheets["Compare"]

                    fmt_part1 = workbook.add_format({"bg_color": "#C6EFCE"})  # 绿：完全匹配
                    fmt_part2 = workbook.add_format({"bg_color": "#F8CBAD"})  # 橙：数量不一致
                    fmt_part3 = workbook.add_format({"bg_color": "#FFEB9C"})  # 黄：只在 A
                    fmt_part4 = workbook.add_format({"bg_color": "#FFC7CE"})  # 红：只在 B
                    fmt_part5 = workbook.add_format({"bg_color": "#D9E1F2"})  # 蓝：无 PO
                    fmt_total = workbook.add_format({"bold": True})

                    cat_col_idx = df_final.columns.get_loc("Category")
                    for row_idx in range(1, len(df_final) + 1):
                        cat = df_final.iloc[row_idx - 1, cat_col_idx]
                        fmt = None
                        if cat.startswith("1. "):
                            fmt = fmt_part1
                        elif cat.startswith("2. Qty"):
                            fmt = fmt_part2
                        elif cat.startswith("3. Only in A"):
                            fmt = fmt_part3
                        elif cat.startswith("4. Only in B"):
                            fmt = fmt_part4
                        elif cat.startswith("5. No PO"):
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
        pass

# ========== TOOL 3: Order Merge Tool V2 ==========
elif tool == "Order Merge Tool V2":
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

    import streamlit.components.v1 as components
    BUILD_ID = "OMTV2-2025-10-15-02"

    file = st.file_uploader("Upload the Excel file (old layout)", type=["xlsx","xls"], key="order_merge_v2")

    RE_PO = re.compile(r'(?:PO:|<strong>PO:</strong>)\s*#?\s*(\d+)', re.IGNORECASE)
    RE_WS = re.compile(r'\s+')

    REQUIRED_COLS = [
        "DateCreated","OrderNumber","OrderStatus","Product_Description","Size","Colour",
        "CustomerName","Phone","Mobile","DeliveryMode","PublicComments","qtyRequired","SourceFrom"
    ]

    # ----------------- 暴力日期解析：只看第一个 a/b/c，然后重排为 yyyy/mm/dd -----------------
# 1) 暴力日期解析（保持不变）
    RE_DMY = re.compile(r'(\d{1,4})\s*/\s*(\d{1,2})\s*/\s*(\d{2,4})')

    def brutal_extract_ymd(value):
        s = str(value).strip()
        m = RE_DMY.search(s)
        if not m:
            return None
        a, b, c = m.groups()   # a=day, b=month, c=year（原始是 dd/mm/yyyy）
        day = int(a)
        month = int(b)
        year = int(c) if len(c) == 4 else int("20" + c)
        return (year, month, day)   # 返回 (yyyy, mm, dd)

    # 2) 只在“不是 yyyy/mm/dd”时才重排
    RE_YMD_FINAL = re.compile(r'^\s*\d{4}/\d{1,2}/\d{1,2}\s*$')

    def brutal_format_ymd(value):
        s = str(value).strip()
        # 已经是 yyyy/mm/dd → 直接返回，避免二次重排
        if RE_YMD_FINAL.match(s):
            return s
        t = brutal_extract_ymd(s)
        if not t:
            return s if value is not None else ""
        y, m, d = t
        return f"{y}/{m}/{d}"   # 不补零：2025/10/5

    def brutal_min_date(series):
        tuples = [brutal_extract_ymd(v) for v in series]
        tuples = [t for t in tuples if t is not None]
        if not tuples:
            # 没抓到任何 a/b/c，就回退输出首个非空原文（或空）
            for v in series:
                if str(v).strip():
                    return brutal_format_ymd(v)
            return ""
        y, m, d = min(tuples)           # 按 (yyyy, mm, dd) 取最早
        return f"{y}/{m}/{d}"


    # ----------------- 其它工具函数 -----------------
    def clean_str(s):
        if pd.isna(s):
            return ""
        s = str(s)
        s = re.sub(r"<[^>]*>", "", s)            # 去 HTML 标签
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
        """提取 POxxxx；对 '2 Available' 类不处理。"""
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

    # ----------------- 主整合逻辑 -----------------
    def consolidate(df: pd.DataFrame) -> pd.DataFrame:
        # 确保缺失列存在
        for col in REQUIRED_COLS:
            if col not in df.columns:
                df[col] = pd.NA

        # 预处理
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

            # 第4列：HomeDelivery → 任一行为 'home' 则置 1，否则 'pickup'
            delivery_vals = [str(x).strip().lower() for x in g["DeliveryMode"].tolist()]
            home_flag = 1 if any(x == "home" for x in delivery_vals) else "pickup"

            # 第12列：AwaitingPayment 标记
            status_vals = [str(x).strip() for x in g["OrderStatus"].tolist() if str(x).strip()]
            awaiting_flag = 1 if any(x.lower() == "awaiting payment" for x in status_vals) else ""

            # 第15列：Items
            items = [x for x in g["_ItemLine"].tolist() if x]
            items_text = ", ".join(items)

            # 第14列：POs（扁平+去重）
            po_list, seen_po = [], set()
            for sub in g["_POs"].tolist():
                for x in sub:
                    if x not in seen_po:
                        seen_po.add(x)
                        po_list.append(x)
            po_text = ", ".join(po_list)

            # 第7列：ContactPhones（去重）
            phone_opts = [x for x in g["_Phones"].tolist() if x]
            seen_ph, phone_unique = set(), []
            for x in phone_opts:
                if x not in seen_ph:
                    seen_ph.add(x)
                    phone_unique.append(x)
            phones_text = ", ".join(phone_unique)

            # 第13列：PublicComments（去重）
            comments_vals = [clean_str(x) for x in g["PublicComments"].tolist() if clean_str(x)]
            seen_c, comments_unique = set(), []
            for x in comments_vals:
                if x not in seen_c:
                    seen_c.add(x)
                    comments_unique.append(x)
            comments_text = " | ".join(comments_unique)

            # 第2列：DateCreated（按斜杠三段暴力重排，并在组内取最早一个）→ yyyy/mm/dd
            date_value = brutal_min_date(g["DateCreated"])

            # 第6列：CustomerName（首个非空）
            customer = first_nonempty(g["CustomerName"].tolist())

            row = {
                "OrderNumber": order,             # 1
                "DateCreated": date_value,        # 2
                "Col3": "",                       # 3
                "HomeDelivery": home_flag,        # 4
                "Col5": "",                       # 5
                "CustomerName": customer,         # 6
                "ContactPhones": phones_text,     # 7
                "Col8": "", "Col9": "", "Col10": "", "Col11": "",  # 8~11 空
                "AwaitingPayment": awaiting_flag, # 12
                "PublicComments": comments_text,  # 13
                "POs": po_text,                   # 14
                "Items": items_text,              # 15
            }
            rows.append(row)

        out = pd.DataFrame(rows)

        # 确保所有列都存在
        for h in [
            "OrderNumber","DateCreated","Col3","HomeDelivery","Col5","CustomerName","ContactPhones",
            "Col8","Col9","Col10","Col11","AwaitingPayment","PublicComments","POs","Items"
        ]:
            if h not in out.columns:
                out[h] = ""

        # 保险：再统一用暴力格式器
        out["DateCreated"] = out["DateCreated"].astype(str).apply(brutal_format_ymd)

        # 列位重排
        out = out[
            ["OrderNumber","DateCreated","Col3","HomeDelivery","Col5",
             "CustomerName","ContactPhones","Col8","Col9","Col10","Col11",
             "AwaitingPayment","PublicComments","POs","Items"]
        ]
        return out

    def validate_columns(df: pd.DataFrame):
        return [c for c in REQUIRED_COLS if c not in df.columns]

    if file:
        try:
            # 读取：自动兼容 .xlsx / .xls / HTML伪Excel / 误扩展CSV/TSV
            raw_df, converted = read_excel_any(file, return_converted_bytes=True)

            # 若自动发生了格式转换，给出提示与下载按钮
            if converted:
                st.info("🔁 检测到 HTML/CSV 伪装的 Excel，已自动转换为真实 .xlsx。")
                st.download_button(
                    "📥 下载自动转换的 .xlsx",
                    converted,
                    file_name="converted.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # 列校验
            missing = validate_columns(raw_df)
            if missing:
                st.error("❌ 缺少以下必要列，请在原表中补齐后再上传：\n\n- " + "\n- ".join(missing))
            else:
                with st.spinner("Processing…"):
                    merged = consolidate(raw_df)

                # 预览（显示前 50 行）
                st.success(f"✅ 处理完成，共 {len(merged)} 条订单（每个 OrderNumber 一行）。")
                st.dataframe(merged.head(50), use_container_width=True)

                # —— 一键复制（不含表头，复制整表） ——
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

                # 下载结果（Excel）
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
    pass




# ========== TOOL 3: Profit Calculator ==========
elif tool == "Profit Calculator":
    st.subheader("💰 Profit Calculator")
    st.caption("All data is calculated locally · Multi-product supported · Updated by Andy Wang")

    # 产品数量
    num_products = st.number_input("Number of products", min_value=1, max_value=20, value=1)
    cols = st.columns(num_products)

    # 每个产品输入不含税成本 & 服务费百分比
    base_costs = []
    service_rates = []
    for i in range(num_products):
        st.markdown(f"**Product {i+1}**")
        col1, col2 = st.columns([2, 1])
        with col1:
            cost = st.number_input(
                f"Base cost (excl. GST) – Product {i+1}",
                min_value=0.0,
                step=0.01,
                format="%.2f",
                value=None,
                placeholder="E.g. 289.75",
                key=f"base_cost_{i}"
            ) or 0.0
        with col2:
            rate = st.radio(
                f"Service Fee – P{i+1}",
                ["15%", "5%"],
                horizontal=True,
                key=f"rate_{i}"
            )
            rate_val = 0.15 if rate == "15%" else 0.05

        base_costs.append(cost)
        service_rates.append(rate_val)

    # 计算总成本（含服务费和GST）
    total_base_cost = sum(base_costs)
    service_fees = [c * r for c, r in zip(base_costs, service_rates)]
    cost_with_service = [c + s for c, s in zip(base_costs, service_fees)]
    total_cost_excl_gst = sum(cost_with_service)
    total_cost_incl_gst = total_cost_excl_gst * 1.15

    # 运费体积输入
    total_volume = st.number_input("Total volume (m³)", min_value=0.0, step=0.0001, format="%.3f", placeholder="E.g. 0.75") or 0.0
    shipping_unit_price = st.number_input("Shipping unit price (NZD/m³, excl. GST)", min_value=0.0, step=0.01, format="%.2f", value=150.0)
    shipping_cost_excl_gst = total_volume * shipping_unit_price
    shipping_cost_incl_gst = shipping_cost_excl_gst * 1.15

    # 售价输入（含税）
    sale_price = st.number_input("Input sale price (GST included, NZD)", min_value=0.0, step=0.01, format="%.2f", value=None, placeholder="E.g. 1200") or 0.0

    # 其他成本
    rent = sale_price * 0.10

    # 汇总
    total_expense = total_cost_incl_gst + shipping_cost_incl_gst + rent
    profit_with_gst = sale_price - total_expense
    profit_no_gst = profit_with_gst / 1.15 if profit_with_gst else 0.0

    def pct(n):
        return f"{(n/(sale_price or 1)*100):.2f}%" if sale_price else "-"

    # 输出表格
    result_rows = [
        ["Base Cost (Sum)", total_base_cost, pct(total_base_cost)],
        ["Service Fees", sum(service_fees), pct(sum(service_fees))],
        ["Product Total Cost (incl. GST)", total_cost_incl_gst, pct(total_cost_incl_gst)],
        ["Shipping (incl. GST)", shipping_cost_incl_gst, pct(shipping_cost_incl_gst)],
        ["Rent (10%)", rent, pct(rent)],
        ["Total Cost", total_expense, pct(total_expense)],
        ["Profit (incl. GST)", profit_with_gst, pct(profit_with_gst)],
        ["Profit (excl. GST)", profit_no_gst, ""]
    ]
    df_res = pd.DataFrame(result_rows, columns=["Item", "Amount (NZD)", "Ratio to Sale Price"])
    df_res["Amount (NZD)"] = df_res["Amount (NZD)"].map(lambda x: f"{x:.2f}")
    st.table(df_res)

    # 明细说明
    with st.expander("Calculation details"):
        st.markdown(f"""
- **Base Cost Total** = {total_base_cost:.2f} NZD
- **Service Fee Total** = {sum(service_fees):.2f} NZD
- **Cost w/ Service Fee (excl. GST)** = {total_cost_excl_gst:.2f} NZD
- **Cost w/ GST** = {total_cost_incl_gst:.2f} NZD
- **Shipping (excl. GST)** = {shipping_cost_excl_gst:.2f} NZD
- **Shipping (incl. GST)** = {shipping_cost_incl_gst:.2f} NZD
- **Rent** = {rent:.2f} NZD
- **Total Expense** = {total_expense:.2f} NZD
- **Profit (incl. GST)** = {profit_with_gst:.2f} NZD
- **Profit (excl. GST)** = {profit_no_gst:.2f} NZD
        """)

    # 饼图
    if sale_price > 0:
        fig, ax = plt.subplots()
        ax.pie(
            [total_cost_incl_gst, shipping_cost_incl_gst, rent, max(profit_with_gst, 0)],
            labels=["Product Cost", "Shipping", "Rent", "Profit"],
            autopct="%1.1f%%",
            startangle=90
        )
        ax.axis("equal")
        st.pyplot(fig)

    # 导出 Excel
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

# ========== TOOL 4: List Split ==========
elif tool == "List Split":
    st.subheader("📄 List Split")
    st.markdown("Paste copied table data with order number and products. Format: `2*Chair,1*Table`")

    pasted_text = st.text_area("Paste your copied data below (from Excel):")

    if st.button("🔍 Analyze pasted content") and pasted_text:
        try:
            from io import StringIO
            # 读取粘贴的制表符分隔文本；全部按字符串处理，避免数值被转成 1.0 之类
            df_input = pd.read_csv(StringIO(pasted_text), sep="\t", header=None, dtype=str)

            st.write("✅ Preview of parsed input:")
            st.dataframe(df_input, use_container_width=True)

            # 统一清洗：把 None/NaN/'nan' 等转为空串，且去首尾空格
            def _fmt_cell(v):
                if v is None:
                    return ""
                s = str(v).strip()
                return "" if s.lower() in ("nan", "none") else s

            records = []

            for _, row in df_input.iterrows():
                # 订单号：第一列（若只有1列也能取到）
                order_id = _fmt_cell(row.iloc[0]) if len(row) >= 1 else ""

                # 供应商订货号：倒数第二列（需要至少2列才有）
                supplier_code = _fmt_cell(row.iloc[-2]) if len(row) >= 2 else ""

                # 合并成 “供应商订货号//订单号”，若供应商订货号缺失则仅用订单号
                combined_order_ref = f"{supplier_code}//{order_id}" if supplier_code else order_id

                # 产品清单：最后一列，形如 "2*Chair,1*Table"
                product_str = _fmt_cell(row.iloc[-1]) if len(row) >= 1 else ""
                items = [item.strip() for item in product_str.split(',') if '*' in item]

                for item in items:
                    try:
                        qty_str, name = item.split('*', 1)
                        qty_str = _fmt_cell(qty_str)
                        name = _fmt_cell(name)
                        if not name:
                            continue
                        # 兼容 "2" / "2.0" / " 2 " 等
                        qty = int(float(qty_str)) if qty_str else 0
                        records.append({
                            'order': combined_order_ref,  # ← 合并后的 “供应商订货号//订单号”
                            'name': name,
                            'qty': qty
                        })
                    except Exception:
                        st.warning(f"⚠️ Skipped malformed item: {item}")

            if records:
                # 固定列顺序
                df_result = pd.DataFrame(records)[['order', 'name', 'qty']]

                st.info("🧩 已将倒数第二列识别为『供应商订货号』，并与第一列『订单号』合并为：**供应商订货号//订单号**")
                st.success("✅ Processing completed.")
                st.dataframe(df_result, use_container_width=True)

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
