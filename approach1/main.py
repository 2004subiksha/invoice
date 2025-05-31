import os
import re
import json
import pytesseract
import pandas as pd
from pdf2image import convert_from_path
from PIL import Image

pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Subhisha\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

INPUT_DIR = "input"
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def extract_text_and_confidences(image):
    data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
    full_text = "\n".join(data["text"])
    conf_map = {}
    for i in range(len(data["text"])):
        word = data["text"][i].strip()
        conf = int(data["conf"][i])
        if word and conf >= 0:
            conf_map[word] = conf / 100  
    return full_text, conf_map

def extract_text_from_pdf(pdf_path):
    text = ""
    conf_map_all = {}
    images = convert_from_path(pdf_path, dpi=300)
    for idx, image in enumerate(images):
        img_path = os.path.join(OUTPUT_DIR, f"page_{idx}.png")
        image.save(img_path, "PNG")
        ocr_text, conf_map = extract_text_and_confidences(Image.open(img_path))
        text += ocr_text + "\n"
        conf_map_all.update(conf_map)
    return text, conf_map_all

import re

def extract_fields_from_text(text, conf_map):
    def get_conf(val):
        try:
            return round(conf_map.get(val.strip().split()[0], 0), 2)
        except:
            return 0.0

    fields = {}

    # Normalize text for easier parsing
    text = re.sub(r'\n+', '\n', text)
    text = text.replace('$', '')

    # Invoice ID
    invoice_match = re.search(r"INVOICE\s*#\s*(\d+)", text, re.IGNORECASE)
    fields["invoice_id"] = invoice_match.group(1).strip() if invoice_match else ""

    # Date
    date_match = re.search(r"Date:\s*([A-Za-z]+\s+\d{1,2}\s+\d{4})", text)
    fields["date"] = date_match.group(1).strip() if date_match else ""

    # Bill To
    bill_to_match = re.search(r"Bill To:\s*(.*?)\n", text)
    fields["bill_to"] = bill_to_match.group(1).strip() if bill_to_match else ""

    # Ship To (combine all lines under Ship To)
    ship_to_match = re.search(r"Ship To:\s*(.*?)\n(.*?)\n(.*?)\n", text)
    if ship_to_match:
        ship_lines = [ship_to_match.group(i).strip() for i in range(1, 4)]
        fields["ship_to"] = ', '.join(ship_lines)
    else:
        fields["ship_to"] = ""

    # Item name (first non-empty line after "Item" section)
    item_section = re.search(r"Item\s+Quantity\s+Rate\s+Amount\n(.*?)\n", text, re.DOTALL)
    fields["item_name"] = item_section.group(1).strip() if item_section else ""

    # Rate
    rate_match = re.search(r"(\d+)\s+\$?([\d.]+)\s+\$?([\d.]+)", text)
    fields["rate"] = rate_match.group(2).strip() if rate_match else ""

    # Amount
    fields["amount"] = rate_match.group(3).strip() if rate_match else ""

    # Subtotal
    subtotal_match = re.search(r"Subtotal:\s*\$?([\d.]+)", text)
    fields["subtotal"] = subtotal_match.group(1).strip() if subtotal_match else ""

    # Discount
    discount_match = re.search(r"Discount\s*\(\d+%\):\s*\$?([\d.]+)", text)
    fields["discount"] = discount_match.group(1).strip() if discount_match else ""

    # Shipping
    shipping_match = re.search(r"Shipping:\s*\$?([\d.]+)", text)
    fields["shipping"] = shipping_match.group(1).strip() if shipping_match else ""

    # Total
    total_match = re.search(r"Total:\s*\$?([\d.]+)", text)
    fields["total"] = total_match.group(1).strip() if total_match else ""

    # Order ID
    order_id_match = re.search(r"Order ID\s*:\s*(US-[\w-]+)", text)
    fields["order_id"] = order_id_match.group(1).strip() if order_id_match else ""

    return fields



def save_fields_to_json(fields, name):
    json_path = os.path.join(OUTPUT_DIR, f"{name}.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(fields, f, indent=4)
    print(f"[✔] JSON saved to: {json_path}")

def save_fields_to_excel(fields, name):
    df = pd.DataFrame([fields])
    excel_path = os.path.join(OUTPUT_DIR, f"{name}.xlsx")
    df.to_excel(excel_path, index=False)
    print(f"[✔] Excel saved to: {excel_path}")

def process_invoice(pdf_path):
    file_name = os.path.splitext(os.path.basename(pdf_path))[0]
    print(f"[•] Processing: {file_name}")
    text, conf_map = extract_text_from_pdf(pdf_path)
    fields = extract_fields_from_text(text, conf_map)
    save_fields_to_json(fields, file_name)
    save_fields_to_excel(fields, file_name)
    print(f"[✓] Done processing: {file_name}\n")

if __name__ == "__main__":
    for file in os.listdir(INPUT_DIR):
        if file.lower().endswith(".pdf"):
            pdf_path = os.path.join(INPUT_DIR, file)
            process_invoice(pdf_path)

