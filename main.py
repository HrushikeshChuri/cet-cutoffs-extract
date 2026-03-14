import streamlit as st
import os
import random
from parser import extract_pdf_to_excel

INPUT_FOLDER = "input"
OUTPUT_FOLDER = "output"

os.makedirs(INPUT_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

st.title("CET Cutoff PDF → Excel Converter")

uploaded_file = st.file_uploader("Upload CET Cutoff PDF", type=["pdf"])

if uploaded_file:

    filename = uploaded_file.name
    file_path = os.path.join(INPUT_FOLDER, filename)

    # Rename if file exists
    if os.path.exists(file_path):

        name, ext = os.path.splitext(filename)
        random_num = random.randint(1000,9999)
        filename = f"{name}_{random_num}{ext}"
        file_path = os.path.join(INPUT_FOLDER, filename)

    # Save uploaded file
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success(f"File saved: {filename}")

    if st.button("Extract Cutoff Data"):

        with st.spinner("Processing PDF..."):

            output_file = extract_pdf_to_excel(file_path, OUTPUT_FOLDER)

        st.success("Extraction Complete!")

        with open(output_file, "rb") as f:
            st.download_button(
                label="Download Excel File",
                data=f,
                file_name=os.path.basename(output_file),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )