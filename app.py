import os
import tempfile
import streamlit as st
from org_chart_tool import build

st.set_page_config(
    page_title="Org Chart Generator",
    page_icon="🏢",
    layout="wide"
)

st.title("🏢 Excel to PowerPoint Org Chart Generator")

st.write(
    "Upload an Excel workbook and generate a PowerPoint organizational chart."
)

company = st.text_input("Company Name")

uploaded_file = st.file_uploader(
    "Upload Excel Workbook",
    type=["xlsx"]
)

if uploaded_file:

    if st.button("Generate Org Chart"):

        with tempfile.TemporaryDirectory() as temp_dir:

            input_file = os.path.join(temp_dir, uploaded_file.name)

            with open(input_file, "wb") as f:
                f.write(uploaded_file.read())

            output_file = os.path.join(temp_dir, "OrgChart.pptx")

            with st.spinner("Generating PowerPoint..."):

                build(
                    input_file,
                    output_file,
                    company if company else uploaded_file.name.replace(".xlsx", "")
                )

            with open(output_file, "rb") as f:
                st.success("PowerPoint generated successfully!")

                st.download_button(
                    "📥 Download PowerPoint",
                    f,
                    file_name="OrgChart.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )