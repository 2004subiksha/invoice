import os
import re
import json
import pandas as pd
import fitz  # PyMuPDF
import pytesseract
from PIL import Image

# ---- Paths ----
tesseract_path = r"C:\Users\Subhisha\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
pdf_path = r"C:\Users\Subhisha\Desktop\invoice_extractor\approach_2\input\invoice_Aaron Hawkins_37425.pdf"
output_folder = r"C:\Users\Subhisha\Desktop\invoice_extractor\approach_2\output"
raw_txt_path = os.path.join(output_folder, "raw_txt.txt")

pytesseract.pytesseract.tesseract_cmd = tesseract_path
os.makedirs(output_folder, exist_ok=True)

# ---- Step 1: Extract text and word-level confidence ----
def extract_text_with_confidence(pdf_file):
    doc = fitz.open(pdf_file)
    full_text = ""
    word_confidence = {}

    for page_num in range(len(doc)):
        pix = doc[page_num].get_pixmap(dpi=300)
        img_path = os.path.join(output_folder, f"page_{page_num + 1}.png")
        pix.save(img_path)

        # Get raw text
        text = pytesseract.image_to_string(Image.open(img_path))
        full_text += text + "\n"

        # Get word-level confidence
        data = pytesseract.image_to_data(Image.open(img_path), output_type=pytesseract.Output.DICT)
        for i in range(len(data['text'])):
            word = data['text'][i].strip()
            conf = int(data['conf'][i])
            if word and conf > 0:
                word_confidence[word] = conf

    return full_text, word_confidence

raw_text, word_confidence = extract_text_with_confidence(pdf_path)

# Save raw text
with open(raw_txt_path, "w", encoding="utf-8") as f:
    f.write(raw_text)

# ---- Step 2: Extract fields using regex ----
patterns = {
    "Invoice Number": r"#\s*(\d+)",
    "Date": r"Date:\s*([A-Za-z]+\s+\d{1,2}\s+\d{4})",
    "Customer": r"Bill To:\s*(.*)",
    "Ship Mode": r"Ship Mode:\s*(.*)",
    "Balance Due": r"Balance Due:\s*\$([\d.]+)",
    "Subtotal": r"Subtotal:\s*\$([\d.]+)",
    "Discount": r"Discount.*:\s*\$([\d.]+)",
    "Shipping": r"Shipping:\s*\$([\d.]+)",
    "Total": r"Total:\s*\$([\d.]+)",
    "Order ID": r"Order ID\s*:\s*([A-Z0-9\-]+)",
}

extracted_data = {}
for key, pattern in patterns.items():
    match = re.search(pattern, raw_text)
    value = match.group(1).strip() if match else ""
    # Estimate confidence based on most confident word in value
    words = value.split()
    confidence = max([word_confidence.get(w, 0) for w in words], default=0)
    extracted_data[key] = {
        "value": value,
        "confidence": confidence
    }

# ---- Step 3: Add calculated final_total ----
try:
    subtotal = float(extracted_data["Subtotal"]["value"])
    discount = float(extracted_data["Discount"]["value"])
    final_total = round(subtotal - discount, 2)
except ValueError:
    final_total = ""

extracted_data["Final Total"] = {
    "value": str(final_total),
    "confidence": 100  # calculated
}

# ---- Step 4: Save to JSON ----
json_path = os.path.join(output_folder, "output.json")
with open(json_path, "w", encoding="utf-8") as f:
    json.dump(extracted_data, f, indent=4)

# ---- Step 5: Save to Excel ----
flat_data = {k: v["value"] for k, v in extracted_data.items()}
excel_path = os.path.join(output_folder, "output.xlsx")
df = pd.DataFrame([flat_data])
df.to_excel(excel_path, index=False)

print(" JSON + Excel with confidence + final total saved to:", output_folder)
