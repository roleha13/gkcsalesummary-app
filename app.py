# app.py
import os
import tempfile
import streamlit as st
from processor import process_pdfs_to_excel

st.set_page_config(page_title="Golden Key Casino Sales Analyzer", layout="wide")

st.title("GKC – Sales Analysis Report Generator")
st.markdown("Upload DAILY RESULTS SUMMARY PDF files to generate the monthly Excel sales report with visuals.")

uploaded_files = st.file_uploader(
    "Upload DAILY RESULTS SUMMARY PDFs",
    type=["pdf"],
    accept_multiple_files=True,
)

if st.button("Generate Excel Report"):
    if not uploaded_files:
        st.warning("Please upload at least one PDF file.")
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_paths = []
            for uploaded_file in uploaded_files:
                file_path = os.path.join(tmpdir, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.read())
                pdf_paths.append(file_path)

            output_path = process_pdfs_to_excel(pdf_paths, tmpdir)

            with open(output_path, "rb") as f:
                st.success("Excel report generated successfully.")
                st.download_button(
                    label="Download Excel Report",
                    data=f,
                    file_name=os.path.basename(output_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
