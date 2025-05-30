import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io
import re
import os
import pandas as pd
import magic  # For MIME detection

# Path to Tesseract executable (for Windows users)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def extract_images_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        img_list = page.get_images(full=True)
        for img in img_list:
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image = Image.open(io.BytesIO(image_bytes))
            images.append(image)
    return images

def ocr_images(images):
    text = ""
    for img in images:
        if img.mode != "RGB":
            img = img.convert("RGB")
        text += pytesseract.image_to_string(img) + "\n"
    return text

def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = ""
    for page_num in range(len(doc)):
        page = doc[page_num]
        page_text = page.get_text()
        full_text += page_text + "\n"
    images = extract_images_from_pdf(pdf_path)
    if images:
        full_text += "\n" + ocr_images(images)
    return full_text

def read_csv_text(file_path):
    try:
        df = pd.read_csv(file_path, dtype=str, engine='python')
        text = "\n".join([" ".join(row.dropna()) for _, row in df.iterrows()])
        return text
    except Exception as e:
        print(f"Error reading CSV: {e}")
        return ""

def read_txt_text(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print(f"Error reading TXT: {e}")
        return ""

def ocr_image_file(file_path):
    try:
        image = Image.open(file_path)
        return ocr_images([image])
    except Exception as e:
        print(f"Error opening image: {e}")
        return ""

def extract_reservation_number(text):
    keyword_regex = re.compile(
        r'\b(?:confirmation number|confirmation|reservation number|reservation|booking reference|booking|res\s*#?)\b[:\s\-]*',
        re.IGNORECASE
    )
    code_regex = re.compile(r'\b([A-Z0-9]{5,10}[-]?[A-Z0-9]{0,5})\b')

    lines = text.split('\n')
    for i, line in enumerate(lines):
        if keyword_regex.search(line):
            match = code_regex.search(line)
            if match:
                return match.group(1).strip()
            for j in range(1, 4):
                if i + j < len(lines):
                    next_line = lines[i + j].strip()
                    match = code_regex.search(next_line)
                    if match:
                        return match.group(1).strip()
    return None

def extract_email(text):
    email_regex = re.compile(r'[\w\.-]+@[\w\.-]+\.\w+')
    emails = email_regex.findall(text)
    for email in emails:
        if len(email) >= 6 and '.' in email.split('@')[-1]:
            return email.strip()
    return None

def extract_resort_info(text):
    resort_keywords = ['resort', 'hotel', 'club', 'inn', 'lodge', 'villa', 'motel', 'suite']
    exclude_keywords = ['benefits', 'ownerguide', 'deals', 'offers', 'help', 'summary', 'gmail']

    resort_name = None
    for line in text.split('\n'):
        line_lower = line.lower().strip()
        if any(kw in line_lower for kw in resort_keywords):
            if not any(ex_kw in line_lower for ex_kw in exclude_keywords):
                if 3 <= len(line) <= 60 and re.search('[a-zA-Z]', line):
                    resort_name = line.strip()
                    break

    date_pattern = r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{1,2}[ ]?[A-Za-z]{3,9}[, ]+\d{2,4}|\d{4}[/-]\d{1,2}[/-]\d{1,2})'
    dates = re.findall(date_pattern, text)
    check_in = dates[0] if len(dates) > 0 else None
    check_out = dates[1] if len(dates) > 1 else None

    cost_pattern = r'([$€£]\s*\d{1,3}(?:[,\d]{0,12})(?:\.\d{1,2})?)'
    costs = re.findall(cost_pattern, text)
    total_cost = costs[0].replace(' ', '') if costs else None

    reservation_number = extract_reservation_number(text)

    return resort_name, check_in, check_out, total_cost, reservation_number

def main():
    file_path = input("Enter path to the file: ").strip()

    mime = magic.Magic(mime=True)
    mime_type = mime.from_file(file_path)
    print(f"Detected MIME type: {mime_type}")

    text = ""

    if mime_type == 'application/pdf':
        print("📄 Extracting text and OCR from PDF...")
        text = extract_text_from_pdf(file_path)

    elif mime_type.startswith('image/'):
        print(f"🖼️ Performing OCR on image file ({mime_type})...")
        text = ocr_image_file(file_path)

    elif mime_type == 'text/csv' or file_path.lower().endswith('.csv'):
        print("📄 Reading CSV file content...")
        text = read_csv_text(file_path)

    elif mime_type == 'text/plain' or file_path.lower().endswith('.txt'):
        print("📄 Reading TXT file content...")
        text = read_txt_text(file_path)

    else:
        print(f"❌ Unsupported or unknown MIME type: {mime_type}")
        return

    print("\n=== Extracted Text Preview ===")
    print(text[:2000])  # Preview first 2000 characters

    resort_name, check_in, check_out, total_cost, reservation_number = extract_resort_info(text)
    email = extract_email(text)

    print("\n🔍 Extracted Info:")
    print(f"Resort Name: {resort_name or 'Not found'}")
    print(f"Check-in Date: {check_in or 'Not found'}")
    print(f"Check-out Date: {check_out or 'Not found'}")
    print(f"Total Cost: {total_cost or 'Not found'}")
    print(f"Reservation Number: {reservation_number or 'Not found'}")
    print(f"Email: {email or 'Not found'}")

    # ✅ File Validity Check
    is_valid = True
    reasons = []

    if not resort_name or len(resort_name) < 3 or not re.search(r'[a-zA-Z]', resort_name):
        is_valid = False
        reasons.append("Invalid or missing resort name")

    if not reservation_number or not re.match(r'^[A-Z0-9\-]{5,10}$', reservation_number):
        is_valid = False
        reasons.append("Invalid or missing reservation number")

    if not email or not re.match(r'^[\w\.-]+@[\w\.-]+\.\w+$', email):
        is_valid = False
        reasons.append("Invalid or missing email address")

    print("\n✅ File Validity Check:")
    if is_valid:
        print("✅ This file is likely a valid reservation document.")
    else:
        print("❌ This file is NOT valid for reservation processing.")
        for reason in reasons:
            print(f" - {reason}")

if __name__ == "__main__":
    main()
