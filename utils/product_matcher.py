import re
from io import BytesIO
from functools import lru_cache

import pandas as pd
import requests
import streamlit as st
from rapidfuzz import process, fuzz

from utils.constants import PRODUCT_INFO_URL
from utils.excel_io import read_excel_any


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


class ProductMatcher:
    def __init__(self, names_all, cbms_all):
        self.names_all = names_all
        self.cbms_all = cbms_all

        self.product_dict_raw = dict(zip(names_all, cbms_all))

        self.norm_index = {}
        self.fp_index = {}
        self.names_norm_list = []

        for n, c in zip(names_all, cbms_all):
            n_norm = normalize_name(n)
            n_fp = " ".join(sorted(set(n_norm.split())))
            self.norm_index[n_norm] = c
            self.fp_index[n_fp] = c
            self.names_norm_list.append(n_norm)

        # 给实例方法加缓存（避免 self 不可 hash 的问题）
        self._match_cached = lru_cache(maxsize=4096)(self._match_impl)

    def match(self, name: str):
        return self._match_cached(name or "")

    def _match_impl(self, name: str):
        if not name:
            return None

        raw = self.product_dict_raw.get(name)
        if raw is not None:
            return raw

        n_norm = normalize_name(name)
        got = self.norm_index.get(n_norm)
        if got is not None:
            return got

        n_fp = " ".join(sorted(set(n_norm.split())))
        got = self.fp_index.get(n_fp)
        if got is not None:
            return got

        tokens = n_norm.split()
        prefix = " ".join(tokens[:3]) if len(tokens) >= 3 else " ".join(tokens)
        if prefix:
            m_prefix = process.extractOne(
                prefix,
                [" ".join(t.split()[:3]) for t in self.names_norm_list],
                scorer=fuzz.token_set_ratio,
                score_cutoff=90,
            )
            if m_prefix:
                _, _, matched_idx = m_prefix
                return self.cbms_all[matched_idx]

        m1 = process.extractOne(
            n_norm, self.names_norm_list, scorer=fuzz.token_set_ratio, score_cutoff=88
        )
        if m1:
            _, _, matched_idx = m1
            return self.cbms_all[matched_idx]

        m2 = process.extractOne(
            n_norm, self.names_norm_list, scorer=fuzz.partial_ratio, score_cutoff=85
        )
        if m2:
            _, _, matched_idx = m2
            return self.cbms_all[matched_idx]

        return None


@st.cache_resource
def get_product_matcher() -> ProductMatcher:
    resp = requests.get(PRODUCT_INFO_URL, timeout=30)
    resp.raise_for_status()
    df = read_excel_any(BytesIO(resp.content))

    if {"Product Name", "CBM"} - set(df.columns):
        raise ValueError("`Product Name` and `CBM` columns are required in product_info.xlsx")

    names = df["Product Name"].fillna("").astype(str).tolist()
    cbms = pd.to_numeric(df["CBM"], errors="coerce").fillna(0).tolist()
    return ProductMatcher(names, cbms)


def match_cbm(name: str) -> float:
    matcher = get_product_matcher()
    v = matcher.match(name or "")
    return float(v) if v is not None else 0.0
