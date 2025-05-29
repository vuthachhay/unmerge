import streamlit as st
from openpyxl import load_workbook
import io

st.title("Unmerge Excel Cells and Download")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    in_mem_file = io.BytesIO(uploaded_file.read())
    wb = load_workbook(in_mem_file)
    for ws in wb.worksheets:
        merged_ranges = list(ws.merged_cells.ranges)
        for cell_range in merged_ranges:
            ws.unmerge_cells(str(cell_range))
    out_mem_file = io.BytesIO()
    wb.save(out_mem_file)
    out_mem_file.seek(0)
    st.download_button(
        label="Download unmerged Excel file",
        data=out_mem_file,
        file_name="unmerged_" + uploaded_file.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
