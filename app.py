import streamlit as st
import tempfile
import os
from processor import process_excel

st.title("CCTV Excel Processor")

st.write("Upload an inspection Excel file. The tool will clean the data and generate a formatted table image.")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file is not None:

    st.success("File uploaded")

    # create temporary input file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.read())
        input_path = tmp.name

    output_excel = "processed.xlsx"
    output_png = "table.png"

    if st.button("Process file"):

        with st.spinner("Processing..."):
            process_excel(input_path, output_excel, output_png)

        st.success("Processing finished")

        # preview image
        st.image(output_png)

        # download Excel
        with open(output_excel, "rb") as f:
            st.download_button(
                label="Download Excel",
                data=f,
                file_name="processed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # download PNG
        with open(output_png, "rb") as f:
            st.download_button(
                label="Download Image",
                data=f,
                file_name="table.png",
                mime="image/png"
            )
