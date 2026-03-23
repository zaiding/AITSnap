import streamlit as st
import tempfile
import base64
import streamlit.components.v1 as components
from processor import process_excel, process_excel_ai_agent, analyze_data


API_KEY = st.secrets["GEMINI_API_KEY"]
MY_INSTRUCTIONS = (
    "Analyze the following inspection data and produce a concise technical conclusion. "
    "Structure your answer as: "
    "1. General condition (overall assessment) "
    "2. Main defects (type, location, severity, brief interpretation) "
    "3. Impact (structural and/or hydraulic) "
    "4. Risks (short-term / long-term) "
    "5. Recommended actions | "
    "Rules: - Be concise and technical - Synthesize, do not list all data "
    "- Do not invent missing information - Highlight uncertainties if any"
)


st.title("AIT Snap")
st.write("Upload an inspection Excel file from AIT platform, process it, then download the Excel or image output.")

# Session state initialisation
for key, default in [
    ("processed", False),
    ("output_excel", None),
    ("output_png", None),
    ("image_bytes", None),
    ("excel_bytes", None),
    ("raw_data", None),
    ("ai_analysis", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


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
            # Fix: process_excel_ai_agent takes only input_xlsx (+ optional sheet_name)
            process_excel(input_path, output_excel, output_png)
            st.session_state.raw_data = process_excel_ai_agent(input_path)

        with st.spinner("Running AI analysis..."):
            st.session_state.ai_analysis = analyze_data(
                MY_INSTRUCTIONS, st.session_state.raw_data, API_KEY
            )

        with open(output_excel, "rb") as f:
            st.session_state.excel_bytes = f.read()

        with open(output_png, "rb") as f:
            st.session_state.image_bytes = f.read()

        st.session_state.output_excel = output_excel
        st.session_state.output_png = output_png
        st.session_state.processed = True


if st.session_state.processed:
    st.success("Processing finished")

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
            <html>
            <body style="margin:0; padding:0; overflow:hidden; background:transparent;">
                <div style="height:38px; display:flex; align-items:center; justify-content:center;">
                    <button id="copyBtn" onclick="copyImage()" style="
                        width:100%;
                        height:38px;
                        font-size:15px;
                        border-radius:8px;
                        border:1px ridge;
                        border-color : rgba(255, 0, 0, .5);
                        background:transparent;
                        cursor:pointer;
                        margin:0;
                        transition:all 0.2s ease;
                    ">
                        Copy Image
                    </button>
                </div>

                <script>
                let currentTextColor = "inherit";

                function applyTheme() {{
                    try {{
                        const btn = document.getElementById("copyBtn");
                        const parentBody = window.parent.document.body;
                        const parentStyles = window.parent.getComputedStyle(parentBody);

                        currentTextColor = parentStyles.color || "inherit";

                        btn.style.color = currentTextColor;
                        btn.style.background = "transparent";
                        btn.style.borderColor = currentTextColor;
                    }} catch (e) {{}}
                }}

                async function copyImage() {{
                    const response = await fetch("data:image/png;base64,{img_base64}");
                    const blob = await response.blob();
                    await navigator.clipboard.write([
                        new ClipboardItem({{"image/png": blob}})
                    ]);
                }}

                const btn = document.getElementById("copyBtn");

                btn.addEventListener("mouseenter", () => {{
                    btn.style.background = "rgba(127,127,127,0.12)";
                }});

                btn.addEventListener("mouseleave", () => {{
                    btn.style.background = "transparent";
                }});

                applyTheme();
                setInterval(applyTheme, 1000);
                </script>
            </body>
            </html>
            """,
            height=38,
        )

    # AI analysis result displayed in a text area
    st.subheader("AI Analysis")
    st.text_area(
        label="",
        value=st.session_state.ai_analysis or "",
        height=400,
        disabled=True,
    )
