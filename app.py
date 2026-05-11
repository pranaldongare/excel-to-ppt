"""
Streamlit UI: upload an Excel file, generate a PPTX deck (one slide per row).
Run with: streamlit run app.py
"""

from datetime import datetime

import pandas as pd
import streamlit as st

from ppt_builder import DEFAULT_TOP_TITLE, REQUIRED_COLUMNS, build_presentation


st.set_page_config(page_title="Excel → PPT Worklet Generator", layout="wide")

st.title("Excel → PPT Worklet Generator")
st.caption("Upload your Worklets Excel file. One slide will be generated per row.")

with st.expander("Expected Excel columns", expanded=False):
    st.write("Your Excel must have these columns (header names must match exactly):")
    for col in REQUIRED_COLUMNS:
        st.markdown(f"- `{col}`")

top_title = st.text_input("Top headline (appears on every slide)", value=DEFAULT_TOP_TITLE)

uploaded = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded is not None:
    try:
        df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Could not read the Excel file: {e}")
        st.stop()

    df = df.fillna("")
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        st.error(f"Missing required column(s): {', '.join(missing)}")
        st.stop()

    st.success(f"Loaded {len(df)} row(s).")
    st.dataframe(df, use_container_width=True)

    if st.button("Generate PPT", type="primary"):
        with st.spinner("Generating presentation..."):
            pptx_bytes = build_presentation(
                df.to_dict(orient="records"),
                top_title=top_title,
            )
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            label="Download PPT",
            data=pptx_bytes,
            file_name=f"worklets_{ts}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
