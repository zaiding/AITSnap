import streamlit as st
import tempfile
import base64
import streamlit.components.v1 as components
from processor import process_excel, process_excel_ai_agent, analyze_data

API_KEY = st.secrets["GEMINI_API_KEY"]
MAX_ROWS_PER_IMAGE = 25

MY_INSTRUCTIONS = (
            "Write a concise, high-level conclusion of the SewerBall Camera inspection."
            "Requirements:"
            "Use a short narrative format (1–2 paragraphs maximum)"
            "Write in a natural, human-like way, as if an experienced engineer is summarizing their impression"
            "Avoid generic or repetitive AI-style phrasing (e.g., Overall, it can be concluded that...)"
            "Write in a technical basic way with basic english words"
            "Most important point : Vary sentence structure and wording to feel organic and not templated"
            "Focus only on the most important defects and overall condition"
            "Summarize trends (e.g., worsening defects) instead of listing all observations"
            "Highlight the most critical issue and its approximate location, and take in consideration the severity of the defects"
            "Mention other defects briefly without detailing each one"
            "End with a clear overall impression of the pipe condition"
            "Avoid:"
            "Bullet points, lists, or section headings"
            "Breaking the answer into sections"
            "Exhaustive defect-by-defect descriptions"
            "Overly technical or verbose explanations"
            "Generate 3 different versions of the conclusion. Each version should:"
            "Convey the same key information"
            "Use different wording, structure, and tone"
            "Feel like it was written independently"
            "Separate the 3 versions clearly using:"
            "Conclusion 1:"
            "Conclusion 2:"
            "Conclusion 3:"
)


st.title("AIT Snap")
st.write("Upload an inspection Excel file from AIT platform, process it, then download the Excel or image output.")

# Session state initialisation
for key, default in [
    ("processed", False),
    ("output_excel", None),
    ("output_pngs", None),
    ("image_bytes_list", None),
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
            _, output_pngs = process_excel(
                input_path, output_excel, output_png,
                max_rows_per_image=MAX_ROWS_PER_IMAGE,
            )
            st.session_state.raw_data = process_excel_ai_agent(input_path)

        with open(output_excel, "rb") as f:
            st.session_state.excel_bytes = f.read()

        image_bytes_list = []
        for png_path in output_pngs:
            with open(png_path, "rb") as f:
                image_bytes_list.append(f.read())

        st.session_state.image_bytes_list = image_bytes_list
        st.session_state.output_excel = output_excel
        st.session_state.output_pngs = output_pngs
        st.session_state.processed = True
        st.session_state.ai_analysis = None  # reset analysis on new file


if st.session_state.processed:
    st.success("Processing finished")

    image_bytes_list = st.session_state.image_bytes_list
    n_images = len(image_bytes_list)

    if n_images > 1:
        st.caption(f"Table split into {n_images} images ({MAX_ROWS_PER_IMAGE} observations max per image) for readability.")

    for i, img_bytes in enumerate(image_bytes_list):
        if n_images > 1:
            st.markdown(f"**Part {i + 1} of {n_images}**")
        st.image(img_bytes)

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label=f"Download Image{f' (part {i + 1})' if n_images > 1 else ''}",
                data=img_bytes,
                file_name=f"table_part{i + 1}.png" if n_images > 1 else "table.png",
                mime="image/png",
                use_container_width=True,
                key=f"download_img_{i}",
            )
        with col2:
            img_base64 = base64.b64encode(img_bytes).decode()
            components.html(
                f"""
                <html>
                <body style="margin:0; padding:0; overflow:hidden; background:rgb(19,23,32);">
                    <div style="height:38px; display:flex; align-items:center; justify-content:center;">
                        <button id="copyBtn{i}" onclick="copyImage{i}()" style="
                            width:100%;
                            height:38px;
                            font-size:15px;
                            border-radius:8px;
                            border:1px solid;
                            border-color : rgba(204, 204, 204, .3);
                            cursor:pointer;
                            margin:0;
                            transition:all 0.2s ease;
                        ">
                            Copy Image
                        </button>
                    </div>

                    <script>
                    function applyTheme{i}() {{
                        try {{
                            const btn = document.getElementById("copyBtn{i}");
                            const parentBody = window.parent.document.body;
                            const parentStyles = window.parent.getComputedStyle(parentBody);
                            btn.style.color = parentStyles.color || "inherit";
                            btn.style.background = "transparent";
                        }} catch (e) {{}}
                    }}

                    async function copyImage{i}() {{
                        const response = await fetch("data:image/png;base64,{img_base64}");
                        const blob = await response.blob();
                        await navigator.clipboard.write([
                            new ClipboardItem({{"image/png": blob}})
                        ]);
                    }}

                    const btn{i} = document.getElementById("copyBtn{i}");
                    btn{i}.addEventListener("mouseenter", () => {{ btn{i}.style.background = "rgb(38,39,48)"; }});
                    btn{i}.addEventListener("mouseleave", () => {{ btn{i}.style.background = "transparent"; }});

                    applyTheme{i}();
                    setInterval(applyTheme{i}, 1000);
                    </script>
                </body>
                </html>
                """,
                height=38,
            )

        if n_images > 1 and i < n_images - 1:
            st.divider()

    st.download_button(
        label="Download Excel",
        data=st.session_state.excel_bytes,
        file_name="Observations_table.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # --- AI Analysis section (on demand) ---
    st.divider()
    st.subheader("AI Conclusion")

    extra_instructions = st.text_area(
        label="Additional instructions (optional)",
        placeholder="e.g. Write 4 lines conclusion. Be less intense...",
        height=100,
    )

    final_prompt = MY_INSTRUCTIONS
    if extra_instructions.strip():
        final_prompt += f" Additional instructions from user: {extra_instructions.strip()}"

    if st.session_state.ai_analysis is None:
        if st.button("Generate AI Conclusion", type="primary"):
            with st.spinner("Generating conclusion..."):
                st.session_state.ai_analysis = analyze_data(
                    final_prompt, st.session_state.raw_data, API_KEY
                )
            st.rerun()
    else:
        st.text_area(
            label="",
            value=st.session_state.ai_analysis,
            height=400,
            disabled=True,
        )
        if st.session_state.ai_analysis.startswith("AI analysis failed"):
            st.error("The AI request failed. See details above — this is usually a blocked/empty response from the model, not a crash in the app.")
        if st.button("Regenerate", help="Regenerate with current instructions"):
            st.session_state.ai_analysis = None
            st.rerun()
