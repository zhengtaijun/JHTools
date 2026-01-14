import base64
import re
from io import BytesIO

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image
import pytesseract


def render():
    st.subheader("🖼️ Excel Screenshot to Table")
    st.markdown("Paste (Ctrl+V) or drag a screenshot of an Excel table. Supported formats: JPG, PNG")

    uploaded_image = st.file_uploader("Upload Screenshot", type=["jpg", "jpeg", "png"])
    pasted_image_bytes = st.session_state.get("pasted_image", None)

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

    pasted_json = st.query_params.get("pasted_image")
    if pasted_json:
        try:
            imgdata = base64.b64decode(pasted_json[0].split(",")[-1])
            st.session_state.pasted_image = imgdata
            st.query_params.clear()
        except Exception:
            st.warning("Failed to decode pasted image.")

    image = None
    if uploaded_image:
        image = Image.open(uploaded_image)
        st.image(image, caption="Uploaded image", use_column_width=True)
    elif pasted_image_bytes:
        image = Image.open(BytesIO(pasted_image_bytes))
        st.image(image, caption="Pasted image", use_column_width=True)

    if not image:
        st.info("Please upload or paste a screenshot of a table to begin.")
        return

    with st.spinner("Running OCR..."):
        raw_text = pytesseract.image_to_string(image)
        lines = raw_text.strip().split("\n")
        rows = [re.split(r'\t+|\s{2,}', line.strip()) for line in lines if line.strip()]
        max_len = max((len(row) for row in rows), default=0)
        rows = [row + [''] * (max_len - len(row)) for row in rows]
        df = pd.DataFrame(rows)

    st.success("✅ OCR complete. Here's the extracted table:")
    st.dataframe(df)

    out = BytesIO()
    df.to_excel(out, index=False, header=False)
    out.seek(0)
    st.download_button(
        "📥 Download as Excel",
        out,
        file_name="extracted_table.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    tsv_string = '\n'.join(['\t'.join(map(str, row)) for row in df.values.tolist()])
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
