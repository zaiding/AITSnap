import streamlit as st
import tempfile
import base64
import streamlit.components.v1 as components
from processor import process_excel

st.title("AIT Snap")
st.write("Upload an inspection Excel file from AIT platform, process it, then download the Excel or image output.")

if "processed" not in st.session_state:
    st.session_state.processed = False
if "output_excel" not in st.session_state:
    st.session_state.output_excel = None
if "output_png" not in st.session_state:
    st.session_state.output_png = None
if "image_bytes" not in st.session_state:
    st.session_state.image_bytes = None
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file is not None:
    st.success("File uploaded")

    if st.button("Process file"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
            tmp_in.write(uploaded_file.read())
            input_path = tmp_in.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out_xlsx:
            output_excel = tmp_out_xlsx.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_out_png:
            output_png = tmp_out_png.name

        with st.spinner("Processing..."):
            process_excel(input_path, output_excel, output_png)

        with open(output_excel, "rb") as f:
            st.session_state.excel_bytes = f.read()

        with open(output_png, "rb") as f:
            st.session_state.image_bytes = f.read()

        st.session_state.output_excel = output_excel
        st.session_state.output_png = output_png
        st.session_state.processed = True

if st.session_state.processed:
    st.success("Processing finished")

    # convert image to base64
    
    st.image(st.session_state.image_bytes)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.download_button(
            label="Download Excel",
            data=st.session_state.excel_bytes,
            file_name="processed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    
    with col2:
        st.download_button(
            label="Download Image",
            data=st.session_state.image_bytes,
            file_name="table.png",
            mime="image/png",
            use_container_width=True,
        )
    
    with col3:
        img_base64 = base64.b64encode(st.session_state.image_bytes).decode()
    
        components.html(
            f"""
            <button onclick="copyImage()" style="
                width:100%;
                color : white;
                font-size : 13
                padding:0.55rem;
                border-radius:8px;
                border:1px solid rgba(250,250,250,0.2);
                background:transparent;
                cursor:pointer;
            ">
            Copy Image
            </button>
    
            <script>
            async function copyImage() {{
                const response = await fetch("data:image/png;base64,{img_base64}");
                const blob = await response.blob();
                await navigator.clipboard.write([
                    new ClipboardItem({{"image/png": blob}})
                ]);
            }}
            </script>
            """,
            height=50,
        )
