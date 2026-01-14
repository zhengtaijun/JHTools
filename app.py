import streamlit as st
from pathlib import Path
from PIL import Image

from tools.trf_volume import render as trf_render
from tools.freight_compare import render as freight_render
from tools.order_merge_v2 import render as omv2_render
from tools.profit_calculator import render as profit_render
from tools.list_split import render as list_split_render
from tools.image_table_extractor import render as img_ocr_render
from tools.google_sheet_query import render as sheet_render


def main():
    root = Path(__file__).resolve().parent
    favicon_path = root / "assets" / "favicon.png"
    favicon = Image.open(favicon_path) if favicon_path.exists() else "🛠️"

    st.set_page_config(
        page_title="JHCH Tools Suite | Andy Wang",
        layout="centered",
        page_icon=favicon
    )

    st.title("🛠️ Jory Henley CHC – Internal Tools Suite")
    st.caption("© 2025 • App author: **Andy Wang**")

    pages = {
        "TRF Volume Calculator": trf_render,
        "Order Merge Tool": freight_render,
        "Order Merge Tool V2": omv2_render,
        "Profit Calculator": profit_render,
        "List Split": list_split_render,
        "Image Table Extractor": img_ocr_render,
        "Google Sheet Query": sheet_render,
    }

    tool = st.sidebar.radio("🧰 Select a tool:", list(pages.keys()), index=0)
    pages[tool]()


if __name__ == "__main__":
    main()
