import os
import re
import shutil
import pdfplumber
import pandas as pd
import xlrd
import openpyxl
from docx import Document
import psutil
import pytesseract
from pdf2image import convert_from_path
import concurrent.futures
import signal
import fitz  # PyMuPDF
from PIL import Image
import unicodedata
from email.utils import parseaddr

# Cấu hình đường dẫn cho Tesseract OCR và Poppler
pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
poppler_path = r"C:\\Users\\HA\\Downloads\\Release-24.08.0-0\\poppler-24.08.0\\Library\\bin"

stop_processing = False
error_log_file = "error_log.txt"
error_length_log_file = "error_log_length.txt"

BLOCKED_DOMAINS = ["topcv.vn", "vieclam24h.vn", "timviecnhanh.vn", "careerbuilder.vn"]

def log_error(message):
    with open(error_log_file, "a", encoding="utf-8") as f:
        f.write(message + "\n")

def log_length_error(message):
    with open(error_length_log_file, "a", encoding="utf-8") as f:
        f.write(message + "\n")

def normalize_text(text):
    text = unicodedata.normalize('NFKC', text)
    text = re.sub(r'\s+', ' ', text)
    return text

import difflib

def is_valid_email(email):
    name, addr = parseaddr(email)
    if '@' in addr and '.' in addr and len(addr) <= 100:
        domain_part = addr.split('@')[-1]
        local_part = addr.split('@')[0]
        # loại bỏ dấu chấm đầu/cuối
        if local_part.startswith('.') or local_part.endswith('.'):
            return False
        for blocked in BLOCKED_DOMAINS:
            if blocked in domain_part:
                return False
            # nếu giống trên 80% thì cũng loại
            if difflib.SequenceMatcher(None, domain_part, blocked).ratio() > 0.8:
                return False
        return True
    return False

def extract_email(text):
    lines = text.splitlines()
    merged_text = ""
    for i in range(len(lines)):
        line = lines[i].strip()
        if i + 1 < len(lines):
            next_line = lines[i+1].strip()
            if "@" in line and not re.search(r"\.\w+$", line) and re.search(r"^[a-zA-Z0-9]", next_line):
                line += next_line
                lines[i+1] = ""
        merged_text += line + " "

    merged_text = merged_text.replace('＠', '@').replace('[at]', '@').replace('(at)', '@')
    merged_text = merged_text.replace('.con', '.com').replace('.corn', '.com').replace(',com', '.com')
    merged_text = merged_text.replace(' ', '').replace('\n', '').lower()

    common_subs = {
        'gma1l': 'gmail', 'gmali': 'gmail', 'gmai1': 'gmail', 'gnail': 'gmail',
        'gmaıl': 'gmail', 'gmaiI': 'gmail', 'gmall': 'gmail', 'gmai|': 'gmail',
        'gmai!': 'gmail', 'gma1!': 'gmail', 'gma11': 'gmail'
    }
    for wrong, right in common_subs.items():
        merged_text = merged_text.replace(wrong, right)

    email_pattern = r'\b[a-zA-Z0-9._%+-]{1,64}@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,10}\b'
    matches = re.findall(email_pattern, merged_text)

    # ⚠️ Chặn email chứa chuỗi ngẫu nhiên dài không hợp lý hoặc suffix sai
    cleaned_matches = []
    for email in matches:
        if is_valid_email(email) and not re.search(r"[a-z]{15,}", email):
            cleaned_matches.append(email)

    return cleaned_matches[0] if cleaned_matches else None

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        doc = fitz.open(pdf_path)
        if len(doc) > 0:
            text = doc[0].get_text()
    except Exception as e:
        log_error(f"[❌] Lỗi đọc PDF bằng PyMuPDF {pdf_path}: {e}")
    return text


def extract_text_from_pdf_fitz(pdf_path):
    text = ""
    try:
        doc = fitz.open(pdf_path)
        if len(doc) > 0:
            text = doc[0].get_text()
    except Exception as e:
        log_error(f"[❌] Lỗi đọc PDF bằng PyMuPDF {pdf_path}: {e}")
    return text


def clean_ocr_text(text):
    text = text.replace('\n', '').replace(' ', '')
    substitutions = {
        'gma1l': 'gmail', 'gmali': 'gmail', 'gmai1': 'gmail', 'gnail': 'gmail',
        'gmaıl': 'gmail', 'gmaiI': 'gmail', 'gmall': 'gmail', 'gmai|': 'gmail',
        'gmai!': 'gmail', 'gma1!': 'gmail', 'gma11': 'gmail'
    }
    for wrong, right in substitutions.items():
        text = text.replace(wrong, right)
    text = text.replace('©', '@').replace('®', '@').replace('＠', '@') \
               .replace('[at]', '@').replace('(at)', '@') \
               .replace(',com', '.com').replace('.com.', '.com') \
               .replace('.con', '.com').replace('.corn', '.com')
    return text

def extract_text_with_ocr(pdf_path):
    text = ""
    try:
        images = convert_from_path(pdf_path, dpi=800, poppler_path=poppler_path)
        if images:
            raw_text = pytesseract.image_to_string(images[0], lang="eng+jpn")  # chỉ ảnh đầu
            text += clean_ocr_text(raw_text) + "\n"
    except Exception as e:
        log_error(f"[❌] Lỗi OCR PDF {pdf_path}: {e}")
    print(text)
    return text



import zipfile

def is_valid_docx(file_path):
    return zipfile.is_zipfile(file_path)

def extract_text_from_docx(docx_path):
    try:
        doc = Document(docx_path)
        extracted_text = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
        return "\n".join(extracted_text).strip()
    except Exception as e:
        log_error(f"[❌] Lỗi đọc DOCX {docx_path}: {e}")
        return ""



def extract_text_from_excel(excel_path):
    text = ""
    try:
        if excel_path.lower().endswith(".xls"):
            workbook = xlrd.open_workbook(excel_path)
            for sheet in workbook.sheets():
                for row in range(sheet.nrows):
                    text += " ".join(map(str, sheet.row_values(row))) + "\n"
        else:
            df = pd.read_excel(excel_path, engine="openpyxl")
            text += "\n".join(df.astype(str).stack().tolist())
    except Exception as e:
        log_error(f"[❌] Lỗi đọc Excel {excel_path}: {e}")
    return text

def is_file_in_use(file_path):
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            with proc.oneshot():
                if file_path in [f.path for f in proc.open_files()]:
                    return True
        except Exception:
            pass
    return False

def signal_handler(sig, frame):
    global stop_processing
    print("[⚠️] Nhận tín hiệu thoát, dừng xử lý an toàn.")
    stop_processing = True

signal.signal(signal.SIGINT, signal_handler)

def process_file(file_path, cv_folder, unprocessed_folder, complete_folder, error_folder, namelength_folder):
    if stop_processing:
        return

    if is_file_in_use(file_path):
        error_msg = f"[⚠️] File {file_path} đang được sử dụng, bỏ qua."
        print(error_msg)
        log_error(error_msg)
        shutil.move(file_path, os.path.join(unprocessed_folder, os.path.basename(file_path)))
        return

    email = None
    filename = os.path.basename(file_path)

    try:
        if filename.lower().endswith(".pdf"):
            text = extract_text_from_pdf(file_path)
            email = extract_email(text)
            if not email:
                text = extract_text_from_pdf_fitz(file_path)
                email = extract_email(text)
            if not email:
                text = extract_text_with_ocr(file_path)
                email = extract_email(text)
        elif filename.lower().endswith(".docx"):
            text = extract_text_from_docx(file_path)
            email = extract_email(text)
        elif filename.lower().endswith((".xls", ".xlsx")):
            text = extract_text_from_excel(file_path)
            email = extract_email(text)
        else:
            error_msg = f"[⚠️] Bỏ qua file không hỗ trợ: {filename}"
            print(error_msg)
            log_error(error_msg)
            shutil.move(file_path, os.path.join(unprocessed_folder, filename))
            return

        print(f"📧 Email trích xuất từ {filename}: {email}")

        if email:
            new_filename = f"{email}{os.path.splitext(filename)[1]}"
            new_path = os.path.join(cv_folder, new_filename)
            counter = 1
            while os.path.exists(new_path):
                new_filename = f"{email}_{counter}{os.path.splitext(filename)[1]}"
                new_path = os.path.join(cv_folder, new_filename)
                counter += 1

            try:
                shutil.move(file_path, new_path)
                if len(email) > 28:
                    log_length_error(f"{filename} | Email quá dài: {email} ({len(email)} ký tự)")
                    shutil.move(new_path, os.path.join(namelength_folder, os.path.basename(new_path)))
                    print(f"[📏] Email dài > 30 ký tự, đã chuyển vào NameLength: {filename} -> {new_filename}")
                else:
                    shutil.move(new_path, os.path.join(complete_folder, os.path.basename(new_path)))
                    print(f"[✅] Đã đổi tên và di chuyển vào complete: {filename} -> {new_filename}")
            except Exception as e:
                error_msg = f"[❌] Lỗi khi đổi tên file {filename}: {e}"
                print(error_msg)
                log_error(error_msg)
                shutil.move(file_path, os.path.join(error_folder, filename))
        else:
            error_msg = f"[❌] Không tìm thấy email trong file: {filename}"
            print(error_msg)
            log_error(error_msg)
            shutil.move(file_path, os.path.join(unprocessed_folder, filename))

    except Exception as e:
        error_msg = f"[❌] Lỗi xử lý file {filename}: {e}"
        print(error_msg)
        log_error(error_msg)
        shutil.move(file_path, os.path.join(error_folder, filename))

def rename_cv_files(cv_folder):
    unprocessed_folder = os.path.join(cv_folder, "unprocessed")
    complete_folder = os.path.join(cv_folder, "complete")
    error_folder = os.path.join(cv_folder, "error")
    namelength_folder = os.path.join(cv_folder, "NameLength")

    os.makedirs(unprocessed_folder, exist_ok=True)
    os.makedirs(complete_folder, exist_ok=True)
    os.makedirs(error_folder, exist_ok=True)
    os.makedirs(namelength_folder, exist_ok=True)

    files = [os.path.join(cv_folder, f) for f in os.listdir(cv_folder) if os.path.isfile(os.path.join(cv_folder, f))]
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        future_to_file = {
            executor.submit(process_file, file, cv_folder, unprocessed_folder, complete_folder, error_folder, namelength_folder): file
            for file in files
        }
        try:
            for future in concurrent.futures.as_completed(future_to_file):
                if stop_processing:
                    executor.shutdown(wait=False)
                    print("[⚠️] Dừng tất cả tiến trình xử lý.")
                    break
        except Exception as e:
            error_msg = f"[❌] Lỗi trong quá trình xử lý: {e}"
            print(error_msg)
            log_error(error_msg)

if __name__ == "__main__":
    cv_folder = "./result"
    rename_cv_files(cv_folder)
