import streamlit as st
import os
import json
from invoice_pipeline import process_invoice

# --- Page setup ---
st.set_page_config(page_title="Invoice Extractor", layout="centered")
st.title(" Invoice Extractor with OCR, Confidence & Final Total")

st.write("Upload a scanned invoice PDF to extract structured fields along with confidence levels and a computed final total.")

# --- File uploader ---
uploaded_file = st.file_uploader(" Upload Invoice PDF", type=["pdf"])

if uploaded_file:
    # Ensure 'input' directory exists
    os.makedirs("input", exist_ok=True)

    # Save the uploaded PDF locally
    input_path = os.path.join("input", uploaded_file.name)
    with open(input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success("File uploaded successfully!")

    # --- Process invoice ---
    with st.spinner(" Processing invoice..."):
        extracted_json, excel_path = process_invoice(input_path)

    st.success(" Invoice processed successfully!")

    # --- Display output ---
    st.subheader(" Extracted Data")
    for key, obj in extracted_json.items():
        st.markdown(f"**{key}:** {obj['value']} &nbsp; _(Confidence: {obj['confidence']}%)_")

    # --- Download section ---
    st.subheader("⬇ Download Outputs")

    # JSON download
    json_str = json.dumps(extracted_json, indent=2)
    st.download_button("Download JSON", data=json_str, file_name="output.json", mime="application/json")

    # Excel download
    with open(excel_path, "rb") as f:
        st.download_button("Download Excel", data=f, file_name="output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
