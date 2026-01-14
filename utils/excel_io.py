import pandas as pd
from io import BytesIO


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
    sniff = data[:2048].lower()

    def as_bio():
        return BytesIO(data)

    # 1) HTML 伪装 Excel
    if (sniff.lstrip().startswith(b"<html")
        or sniff.lstrip().startswith(b"<!doctype html")
        or b"<table" in sniff):
        tables = pd.read_html(as_bio(), header=None)
        if not tables:
            raise RuntimeError("HTML 文件中未发现可解析的表格。请导出为真正的 Excel。")
        df = tables[0]

        expected_cols = {
            "datecreated","ordernumber","orderstatus","product_description","size",
            "colour","customername","phone","mobile","deliverymode",
            "publiccomments","qtyrequired","sourcefrom"
        }
        first_row = [str(x).strip() for x in df.iloc[0].tolist()]
        if any(x.lower() in expected_cols for x in first_row):
            df.columns = df.iloc[0]
            df = df.drop(df.index[0]).reset_index(drop=True)

        df = df.applymap(lambda x: "" if pd.isna(x) else str(x))

        conv = _to_xlsx_bytes(df) if return_converted_bytes else None
        return (df, conv) if return_converted_bytes else df

    # 2) 真 .xlsx
    if head.startswith(b"PK\x03\x04"):
        try:
            df = pd.read_excel(as_bio(), engine="openpyxl", **kwargs)
        except Exception:
            df = pd.read_excel(as_bio(), **kwargs)
        return (df, None) if return_converted_bytes else df

    # 3) 真 .xls
    if head.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1") or name.endswith(".xls"):
        _ensure_xlrd_ok()
        df = pd.read_excel(as_bio(), engine="xlrd", **kwargs)
        return (df, None) if return_converted_bytes else df

    # 4) CSV/TSV 误扩展
    text_sample = data[:4096].decode("utf-8", errors="ignore")
    if ("\t" in text_sample or "," in text_sample) and ("\n" in text_sample or "\r" in text_sample):
        sep = "\t" if text_sample.count("\t") >= text_sample.count(",") else ","
        df = pd.read_csv(BytesIO(data), sep=sep)
        df = df.applymap(lambda x: "" if pd.isna(x) else str(x))
        conv = _to_xlsx_bytes(df) if return_converted_bytes else None
        return (df, conv) if return_converted_bytes else df

    # 5) 兜底
    df = pd.read_excel(as_bio(), **kwargs)
    return (df, None) if return_converted_bytes else df
