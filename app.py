import streamlit as st
import openpyxl
import pypdf
import zipfile
import os
import gc
import re
import io
import pytesseract
import pandas as pd
import tempfile
from io import BytesIO
from pyzbar.pyzbar import decode
from pdf2image import convert_from_bytes
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- CẤU HÌNH GIAO DIỆN ---
st.set_page_config(page_title="POD Tool", page_icon="⚡", layout="wide")
st.markdown("""<style>.stButton>button { width: 100%; }</style>""", unsafe_allow_html=True)

# --- HÀM HỖ TRỢ ---
def get_letter_code(index):
    result = ""
    while index >= 0:
        result = chr(65 + (index % 26)) + result
        index = (index // 26) - 1
    return result

def sanitize_filename(name):
    clean_name = re.sub(r'[\\/*?:"<>|]', '_', str(name).strip())
    return clean_name.replace(' ', '_') if clean_name else "FILE_UNNAMED"

# --- XỬ LÝ ZIP (ĐỌC TỪ FILE TẠM) ---
def build_zip_final(records, is_opt2=False):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zipf:
        wb = openpyxl.Workbook()
        ws = wb.active
        
        if not is_opt2:
            headers = ["STT", "STT Trang", "Tên File / Mã Vạch", "Status"]
            for i, h in enumerate(headers, 1): ws.cell(1, i, h)
            for idx, rec in enumerate(records, 1):
                ws.cell(idx+1, 1, rec["stt"]); ws.cell(idx+1, 2, rec["stt_trang"])
                ws.cell(idx+1, 3, rec["code"]); ws.cell(idx+1, 4, rec["status"])
            zipf.writestr("DANH_SACH_TONG_HOP.xlsx", save_excel(wb))
        else:
            headers = ["STT", "File Gốc", "STT Trang", "Mã", "Loại", "Trạng Thái"]
            for i, h in enumerate(headers, 1): ws.cell(1, i, h)
            for idx, rec in enumerate(records, 1):
                ws.cell(idx+1, 1, rec["stt"]); ws.cell(idx+1, 2, rec["orig_file"])
                ws.cell(idx+1, 3, rec["stt_trang"]); ws.cell(idx+1, 4, rec["code"])
                ws.cell(idx+1, 5, rec["doc_type"]); ws.cell(idx+1, 6, rec["status"])
            zipf.writestr("DANH_SACH_BSH.xlsx", save_excel(wb))

        for rec in records:
            if rec.get("temp_path") and os.path.exists(rec["temp_path"]):
                zipf.write(rec["temp_path"], f"PDF_DATA/{sanitize_filename(rec['code'])}.pdf")
    
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

def save_excel(wb):
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

# --- GIAO DIỆN ---
option = st.sidebar.radio("CHỌN CÔNG CỤ:", ["1. Quét Mã Vạch & Tách PDF", "2. BSH_OCR & Tách PDF"])

if "records" not in st.session_state: st.session_state.records = []

if st.sidebar.button("Reset Session"):
    for r in st.session_state.records:
        if "temp_path" in r and os.path.exists(r["temp_path"]): os.remove(r["temp_path"])
    st.session_state.records = []
    st.rerun()

# --- LOGIC XỬ LÝ (DÙNG TEMPFILE) ---
def process_page(pdf_bytes, i, is_ocr=False):
    page_num = i + 1
    reader = pypdf.PdfReader(BytesIO(pdf_bytes))
    writer = pypdf.PdfWriter()
    writer.add_page(reader.pages[i])
    
    # Lưu file tạm thay vì BytesIO
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    writer.write(tmp.name)
    
    # OCR/Barcode logic ở đây (đã rút gọn để đảm bảo code chạy)
    img = convert_from_bytes(pdf_bytes, dpi=150, first_page=page_num, last_page=page_num)[0]
    # ... (Giữ nguyên logic OCR cũ của bạn tại đây) ...
    
    return {"temp_path": tmp.name, "preview_img": img}

# --- MAIN ---
if option.startswith("1"):
    files = st.file_uploader("Upload PDF", accept_multiple_files=True)
    if files and st.button("XỬ LÝ"):
        all_recs = []
        for f in files:
            b = f.read()
            for i in range(len(pypdf.PdfReader(BytesIO(b)).pages)):
                res = process_page(b, i)
                all_recs.append({"stt": len(all_recs)+1, "stt_trang": i+1, "code": "A", "status": "OK", **res, "orig_file": f.name})
        st.session_state.records = all_recs
        st.success("Xử lý xong!")

if st.session_state.records:
    if st.button("TẢI ZIP"):
        data = build_zip_final(st.session_state.records)
        st.download_button("DOWNLOAD", data, "KET_QUA.zip", "application/zip")
