"""
Module xử lý lõi: Tự động đối chiếu, gộp file RFI - RP - Bản vẽ và tạo File PDF tổng nối dài.
Dự án: ĐƯỜNG ỐNG DẪN KHÍ LÔ B - Ô MÔN (BLOCK B - O MON GAS PIPELINE PROJECT)
"""

import os
import re
import sys
import shutil
import tempfile
import datetime
import traceback
from typing import List, Dict, Optional, Tuple, Set

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import pymupdf
    HAS_PYMUPDF = True
except ImportError:
    pymupdf = None
    HAS_PYMUPDF = False

try:
    import cv2
    import numpy as np
    HAS_CV2 = True
except ImportError:
    cv2 = None
    np = None
    HAS_CV2 = False

try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    win32com = None
    pythoncom = None
    HAS_WIN32COM = False

try:
    from rapidocr_onnxruntime import RapidOCR
    HAS_RAPIDOCR = True
except ImportError:
    RapidOCR = None
    HAS_RAPIDOCR = False

ocr_engine = None

def get_ocr_engine():
    global ocr_engine
    if ocr_engine is None:
        if RapidOCR is not None:
            ocr_engine = RapidOCR()
    return ocr_engine


def normalize_code(s: str) -> str:
    """Loại bỏ ký tự đặc biệt, dấu cách, sửa lỗi nhận diện OCR phổ biến, chuyển chữ hoa."""
    if not s:
        return ""
    txt = str(s).upper()
    # Sửa lỗi ký tự K dạng CAD hay bị OCR đọc thành <, (, [, {
    txt = re.sub(r'[<{(\[]P', 'KP', txt)
    txt = txt.replace('<', 'K')
    return re.sub(r'[^A-Z0-9]', '', txt)


def extract_seq_number(name: str) -> int:
    """Trích xuất số thứ tự từ tên file/mã báo cáo để sắp xếp từ nhỏ đến lớn (ví dụ CAI-0001 -> 1)."""
    digits = re.findall(r'\d+', name)
    return int(digits[-1]) if digits else 0


def parse_rp_excel(file_path: str) -> Dict:
    """
    Đọc thông tin từ file Báo cáo kiểm tra (RP Excel):
    - Report No
    - RFI No
    - Sheet Bản vẽ (Weld map sheet)
    - Danh sách mối hàn (Joint No)
    - Tên Sheet chứa báo cáo RP
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    
    target_sheet_name = None
    # Ưu tiên sheet báo cáo (bắt đầu bằng RP hoặc chứa mã loại), bỏ qua các sheet Control/Log
    for name in wb.sheetnames:
        uname = name.upper()
        if 'CONTROL' in uname or 'LOG' in uname:
            continue
        if uname.startswith('RP') or 'RP-' in uname:
            target_sheet_name = name
            break
            
    if not target_sheet_name:
        for name in wb.sheetnames:
            uname = name.upper()
            if 'CONTROL' in uname or 'LOG' in uname:
                continue
            target_sheet_name = name
            break
            
    if not target_sheet_name:
        target_sheet_name = wb.sheetnames[0]
        
    ws = wb[target_sheet_name]
    
    report_no = None
    rfi_no = None
    weld_map_no = None
    drawing_sheet_num = None
    joints = []
    
    is_fur = ('FUR' in target_sheet_name.upper()) or ('-FUR-' in os.path.basename(file_path).upper())
    is_icr = ('ICR' in target_sheet_name.upper()) or ('-ICR-' in os.path.basename(file_path).upper())
    is_ppp = ('PPP' in target_sheet_name.upper()) or ('-PPP-' in os.path.basename(file_path).upper())
    is_mrir = ('MRIR' in target_sheet_name.upper()) or ('-MRIR-' in os.path.basename(file_path).upper()) or ('_MRIR' in os.path.basename(file_path).upper()) or ('RP-MRIR' in os.path.basename(file_path).upper()) or ('RP -MRIR' in os.path.basename(file_path).upper()) or os.path.basename(file_path).upper().startswith('MRIR')
    is_dic = ('DIC' in target_sheet_name.upper()) or ('-DIC-' in os.path.basename(file_path).upper()) or ('_DIC' in os.path.basename(file_path).upper()) or ('RP-DIC' in os.path.basename(file_path).upper()) or ('RP -DIC' in os.path.basename(file_path).upper()) or os.path.basename(file_path).upper().startswith('DIC')

    # Đọc thông tin header chung (Report No, RFI No, Bản vẽ số / Weld map No)
    for r in range(1, min(35, ws.max_row + 1)):
        for c in range(1, min(20, ws.max_column + 1)):
            val = ws.cell(r, c).value
            if val is None:
                continue
            sval = str(val).strip()
            
            # Report No
            if not report_no and ('Báo cáo số' in sval or 'Report No.:' in sval or 'Report No:' in sval or (('Báo cáo' in sval or 'Report No' in sval) and ('MRIR' not in sval or is_mrir) and 'KIỂM TRA' not in sval.upper())):
                for off_c in range(1, 5):
                    v = ws.cell(r, c + off_c).value
                    if v and str(v).strip() and not any(k in str(v) for k in (['Phương pháp', 'Cleaning', 'Date'] if is_mrir else ['Phương pháp', 'Cleaning', 'MRIR', 'Date'])):
                        report_no = str(v).strip()
                        break
                if not report_no:
                    v_next = ws.cell(r + 1, c).value
                    if v_next and str(v_next).strip():
                        report_no = str(v_next).strip()
            elif sval.startswith('SWG-') and 'RFI' not in sval and ('MRIR' not in sval or is_mrir) and 'WPS' not in sval and 'WMP' not in sval and 'DAG' not in sval:
                if not report_no or not report_no.startswith('SWG-'):
                    report_no = sval

            # RFI No
            if not rfi_no and ('RFI No.:' in sval or 'Thư mời kiểm tra' in sval or 'RFI No' in sval):
                for off_c in range(1, 5):
                    v = ws.cell(r, c + off_c).value
                    if v and str(v).strip() and ('MRIR' not in str(v) or is_mrir):
                        rfi_no = str(v).strip()
                        break
                if not rfi_no:
                    v_next = ws.cell(r + 1, c).value
                    if v_next and str(v_next).strip() and ('MRIR' not in str(v_next) or is_mrir):
                        rfi_no = str(v_next).strip()
            elif 'RFI-' in sval and ('MRIR' not in sval or is_mrir):
                if not rfi_no or not rfi_no.startswith('SWG-'):
                    rfi_no = sval

            # Weld Map No / Drawing No
            if 'Bản vẽ mối hàn' in sval or 'Weld map' in sval or 'Drawing No' in sval:
                if 'SHT' in sval.upper() or 'SHEET' in sval.upper() or 'WMP' in sval.upper() or 'DAG' in sval.upper():
                    weld_map_no = sval
                else:
                    for off_c in range(1, 5):
                        v = ws.cell(r, c + off_c).value
                        if v and str(v).strip():
                            weld_map_no = str(v).strip()
                            break
            elif ('SHT ' in sval or 'SHEET ' in sval) and not weld_map_no:
                weld_map_no = sval

    # Trích xuất danh sách số mối hàn (Joint No) hoặc danh sách số ống (Pipe No) hoặc vật tư (MRIR):
    if is_icr:
        # Đối với mẫu báo cáo ICR: không sử dụng mối hàn mà sử dụng số lượng ống (Cột F: Ống số / Pipe No.)
        for r in range(1, ws.max_row + 1):
            val_f = ws.cell(r, 6).value # Cột F: Ống số / Pipe No.
            stt_val = ws.cell(r, 1).value # Cột A: STT
            if val_f and str(val_f).strip():
                s_f = str(val_f).strip()
                if s_f.startswith('181-') or (stt_val and str(stt_val).strip().isdigit()):
                    if 'Pipe No' not in s_f and 'Ống số' not in s_f and s_f not in joints:
                        joints.append(s_f)
    elif is_mrir:
        # Đối với mẫu báo cáo MRIR (kiểm tra nhận vật tư)
        start_row = 8
        for r in range(start_row, min(50, ws.max_row + 1)):
            val_a = ws.cell(r, 1).value # STT
            val_e = ws.cell(r, 5).value # Description
            val_h = ws.cell(r, 8).value # Type of Material
            val_j = ws.cell(r, 10).value # Size 1
            if val_a and str(val_a).strip().isdigit() and ((val_e and str(val_e).strip()) or (val_h and str(val_h).strip())):
                desc_parts = []
                if val_h and str(val_h).strip():
                    desc_parts.append(str(val_h).strip())
                elif val_e and str(val_e).strip():
                    desc_parts.append(str(val_e).strip())
                if val_j and str(val_j).strip() and str(val_j).strip() != 'N/A':
                    desc_parts.append(str(val_j).strip())
                item_str = f"Mục {val_a} ({' - '.join(desc_parts)})"
                if item_str not in joints:
                    joints.append(item_str)
    elif is_dic:
        # Đối với mẫu báo cáo DIC (kiểm tra hiệu chuẩn thiết bị hàng ngày)
        start_row = 12
        for r in range(start_row, min(20, ws.max_row + 1)):
            val_type = ws.cell(r, 1).value # Loại Instrument
            val_id = ws.cell(r, 4).value   # I/D
            if val_type and str(val_type).strip():
                stype = str(val_type).strip()
                sid = str(val_id).strip() if val_id else ''
                item_str = f"{stype} ({sid})" if sid else stype
                if item_str not in joints:
                    joints.append(item_str)
    elif is_fur or is_ppp:
        # Đối với mẫu báo cáo FUR và PPP: thông tin đối chiếu ở Rp-FUR / Rp-PPP là cột số mối hàn (Cột B)
        start_row = 10
        for r in range(1, 25):
            val_b = str(ws.cell(r, 2).value or '').strip()
            if 'Số mối hàn' in val_b or 'Joint No' in val_b or 'Mối số' in val_b:
                start_row = r + 1
                break
                
        for r in range(start_row, min(100, ws.max_row + 1)):
            val = ws.cell(r, 2).value # Cột B
            if val is None:
                continue
            sval = str(val).strip()
            if not sval:
                continue
            # Dừng lại nếu gặp phần chân trang / chữ ký / kết luận
            val_a = str(ws.cell(r, 1).value or '').strip().lower()
            if any(stop_word in sval.lower() or stop_word in val_a for stop_word in [
                'kết luận', 'conclude', 'viết tắt', 'abbreviation', 'kiểm tra', 'inspection',
                'chức danh', 'position', 'chữ ký', 'signature', 'họ và tên', 'full name',
                'chú thích', 'note:'
            ]):
                break
            for part in re.split(r'[;,\n]', sval):
                p = part.strip()
                if p and p not in joints:
                    joints.append(p)
    else:
        # Đối với các mẫu báo cáo khác (CAI, FTI, HTI, DIC, ...)
        for r in range(1, min(35, ws.max_row + 1)):
            for c in range(1, min(20, ws.max_column + 1)):
                val = ws.cell(r, c).value
                if val is None:
                    continue
                sval = str(val).strip()
                if ('Mối số' in sval or 'Joint No' in sval) and not joints:
                    row_joints = []
                    for col_idx in range(c + 1, ws.max_column + 1):
                        jv = ws.cell(r, col_idx).value
                        if jv is not None and str(jv).strip():
                            j_str = str(jv).strip()
                            for part in re.split(r'[;,]', j_str):
                                p = part.strip()
                                if p and re.match(r'^(?:KP\d+[-_][A-Za-z0-9]+|W\d+)', p, re.I) and p not in row_joints:
                                    row_joints.append(p)
                    if row_joints:
                        joints.extend(row_joints)
                    else:
                        for row_idx in range(r + 1, min(r + 30, ws.max_row + 1)):
                            jv = ws.cell(row_idx, c).value
                            if jv is None:
                                continue
                            j_str = str(jv).strip()
                            if not j_str:
                                continue
                            val_a = str(ws.cell(row_idx, 1).value or '').strip().lower()
                            if any(sw in j_str.lower() or sw in val_a for sw in ['kết luận', 'viết tắt', 'kiểm tra', 'chữ ký', 'họ và tên']):
                                break
                            for part in re.split(r'[;,\n]', j_str):
                                p = part.strip()
                                if p and p not in joints:
                                    joints.append(p)
                                
    if weld_map_no:
        m = re.search(r'(?:SHT|SHEET)\s*(\d+)(?:\s*/\s*(\d+))?', weld_map_no, re.IGNORECASE)
        if m:
            drawing_sheet_num = m.group(1)

    if not report_no:
        base = os.path.splitext(os.path.basename(file_path))[0]
        report_no = base

    return {
        'file_path': file_path,
        'sheet_name': target_sheet_name,
        'report_no': report_no,
        'rfi_no': rfi_no,
        'weld_map_no': weld_map_no,
        'drawing_sheet_num': drawing_sheet_num,
        'joints': joints,
        'is_icr': is_icr,
        'is_mrir': is_mrir,
        'is_dic': is_dic,
        'item_type': 'PIPE' if is_icr else ('MATERIAL' if is_mrir else ('EQUIPMENT' if is_dic else 'WELD')),
        'seq_no': extract_seq_number(report_no)
    }


def parse_rfi_excel(file_path: str) -> Dict:
    """
    Đọc thông tin và trích xuất danh sách mối hàn / số lượng ống và metadata từ file RFI Excel.
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    target_sheet_name = wb.sheetnames[0]
    for name in wb.sheetnames:
        if any(k in name.upper() for k in ['RFI', 'CAI', 'YÊU CẦU', 'ICR', 'NDT', 'PPP', 'FUR', 'FTI', 'MRIR', 'DIC']):
            target_sheet_name = name
            break
            
    ws = wb[target_sheet_name]
    
    text_corpus = ""
    rfi_no = None
    for row in ws.iter_rows(max_row=25, values_only=True):
        for cell in row:
            if cell:
                c_str = str(cell).strip()
                text_corpus += " " + c_str
                if not rfi_no and 'RFI-' in c_str:
                    m_rfi = re.search(r'(SWG-[A-Za-z0-9\-]+-RFI-[A-Za-z0-9\-]+-\d+)', c_str)
                    if m_rfi:
                        rfi_no = m_rfi.group(1)

    if not rfi_no:
        base = os.path.splitext(os.path.basename(file_path))[0]
        m_base = re.search(r'(SWG-[A-Za-z0-9\-]+-RFI-[A-Za-z0-9\-]+-\d+)', base)
        if m_base:
            rfi_no = m_base.group(1)
        else:
            rfi_no = base
                
    # Trích xuất các mã mối dạng KP...-W...
    matches = re.findall(r'KP\d+[-_][A-Za-z0-9]+', text_corpus)
    joints_rfi = []
    for m in matches:
        m_clean = m.upper().replace('_', '-')
        if m_clean not in joints_rfi:
            joints_rfi.append(m_clean)
            
    # Trích xuất số lượng ống đối với mẫu ICR
    pipe_count = None
    m_pipe = re.search(r'(?:Số\s*lượng|Quanity|Quantity)\s*:\s*(\d+)', text_corpus, re.IGNORECASE)
    if m_pipe:
        pipe_count = int(m_pipe.group(1))

    is_icr = ('ICR' in target_sheet_name.upper()) or ('-ICR-' in os.path.basename(file_path).upper())
    is_mrir = ('MRIR' in target_sheet_name.upper()) or ('-MRIR-' in os.path.basename(file_path).upper()) or ('_MRIR' in os.path.basename(file_path).upper()) or ('RFI-MRIR' in os.path.basename(file_path).upper())
    is_dic = ('DIC' in target_sheet_name.upper()) or ('-DIC-' in os.path.basename(file_path).upper()) or ('_DIC' in os.path.basename(file_path).upper()) or ('RFI-DIC' in os.path.basename(file_path).upper())

    return {
        'file_path': file_path,
        'sheet_name': target_sheet_name,
        'rfi_no': rfi_no,
        'joints': joints_rfi,
        'pipe_count': pipe_count,
        'is_icr': is_icr,
        'is_mrir': is_mrir,
        'is_dic': is_dic,
        'item_type': 'PIPE' if is_icr else ('MATERIAL' if is_mrir else ('EQUIPMENT' if is_dic else 'WELD')),
        'seq_no': extract_seq_number(rfi_no)
    }


def reconcile_joints(joints_rp: List[str], joints_rfi: List[str], item_type: str = 'WELD', rfi_pipe_count: Optional[int] = None) -> Dict:
    """
    Đối chiếu sự phù hợp giữa RP và RFI:
    - Nếu item_type == 'PIPE' (mẫu ICR): đối chiếu theo số lượng ống.
    - Nếu item_type == 'MATERIAL' (mẫu MRIR): đối chiếu hồ sơ nhận vật tư.
    - Nếu item_type == 'EQUIPMENT' (mẫu DIC): đối chiếu hồ sơ hiệu chuẩn thiết bị.
    - Nếu item_type == 'WELD': đối chiếu theo số lượng và danh sách mã mối hàn.
    - Trả về: is_match, status ('PHÙ HỢP' / 'CẢNH BÁO'), chi tiết chênh lệch.
    """
    if item_type == 'MATERIAL':
        count_rp = len(joints_rp)
        count_rfi = 1
        return {
            'is_match': True,
            'status': 'PHÙ HỢP',
            'note': f'Kiểm tra nhận vật tư ({count_rp} mục)' if count_rp > 0 else 'Kiểm tra nhận vật tư',
            'count_rp': count_rp,
            'count_rfi': count_rfi,
            'diff_rp_only': [],
            'diff_rfi_only': []
        }

    if item_type == 'EQUIPMENT':
        count_rp = len(joints_rp)
        count_rfi = 1
        return {
            'is_match': True,
            'status': 'PHÙ HỢP',
            'note': f'Hiệu chuẩn thiết bị ({count_rp} thiết bị)' if count_rp > 0 else 'Hiệu chuẩn thiết bị',
            'count_rp': count_rp,
            'count_rfi': count_rfi,
            'diff_rp_only': [],
            'diff_rfi_only': []
        }

    if item_type == 'PIPE':
        count_rp = len(joints_rp)
        count_rfi = rfi_pipe_count if rfi_pipe_count is not None else len(joints_rfi)
        if count_rp == count_rfi and count_rp > 0:
            return {
                'is_match': True,
                'status': 'PHÙ HỢP',
                'note': f'Khớp 100% số lượng ống ({count_rp} ống)',
                'count_rp': count_rp,
                'count_rfi': count_rfi,
                'diff_rp_only': [],
                'diff_rfi_only': []
            }
        else:
            return {
                'is_match': False,
                'status': 'CẢNH BÁO: LỆCH SỐ LƯỢNG ỐNG',
                'note': f'Lệch số lượng ống (RP: {count_rp} ống, RFI: {count_rfi} ống)',
                'count_rp': count_rp,
                'count_rfi': count_rfi,
                'diff_rp_only': [],
                'diff_rfi_only': []
            }

    norm_rp = {normalize_code(j): j for j in joints_rp}
    norm_rfi = {normalize_code(j): j for j in joints_rfi}
    
    set_rp = set(norm_rp.keys())
    set_rfi = set(norm_rfi.keys())
    
    count_rp = len(joints_rp)
    count_rfi = len(joints_rfi)
    
    diff_rp_only = [norm_rp[k] for k in (set_rp - set_rfi)]
    diff_rfi_only = [norm_rfi[k] for k in (set_rfi - set_rp)]
    
    if set_rp == set_rfi and count_rp == count_rfi:
        return {
            'is_match': True,
            'status': 'PHÙ HỢP',
            'note': 'Khớp 100% số lượng & tên mối',
            'count_rp': count_rp,
            'count_rfi': count_rfi,
            'diff_rp_only': [],
            'diff_rfi_only': []
        }
    else:
        notes = []
        if count_rp != count_rfi:
            notes.append(f"Lệch số lượng (RP: {count_rp}, RFI: {count_rfi})")
        if diff_rp_only:
            notes.append(f"RP thừa / RFI thiếu: {', '.join(diff_rp_only)}")
        if diff_rfi_only:
            notes.append(f"RFI thừa / RP thiếu: {', '.join(diff_rfi_only)}")
            
        return {
            'is_match': False,
            'status': 'CẢNH BÁO: LỆCH MỐI',
            'note': "; ".join(notes),
            'count_rp': count_rp,
            'count_rfi': count_rfi,
            'diff_rp_only': diff_rp_only,
            'diff_rfi_only': diff_rfi_only
        }


def find_matching_rfi(rp_info: Dict, input_dir: str) -> Optional[Tuple[str, str]]:
    """Tìm file RFI tương ứng trong thư mục INPUT."""
    rfi_code = rp_info.get('rfi_no')
    norm_rfi = normalize_code(rfi_code) if rfi_code else None
    
    excel_files = [f for f in os.listdir(input_dir) if f.lower().endswith(('.xlsx', '.xls', '.xlsm')) and not f.startswith('~$')]
    
    # 1. Khớp theo tên file chứa RFI No
    if norm_rfi:
        for f in excel_files:
            if norm_rfi in normalize_code(f):
                full_p = os.path.join(input_dir, f)
                return full_p, parse_rfi_excel(full_p)['sheet_name']
                
    # 2. Khớp theo nội dung bên trong file
    for f in excel_files:
        full_p = os.path.join(input_dir, f)
        try:
            wb = openpyxl.load_workbook(full_p, read_only=True, data_only=True)
            for sname in wb.sheetnames:
                ws = wb[sname]
                for row in ws.iter_rows(max_row=15, values_only=True):
                    for cell in row:
                        if cell and norm_rfi and norm_rfi in normalize_code(str(cell)):
                            return full_p, sname
        except Exception:
            pass

    # 3. Khớp theo mã báo cáo và mã file RFI tương ứng
    if rp_info.get('report_no'):
        rep_norm = normalize_code(rp_info['report_no'])
        for f in excel_files:
            if 'RFI' in f.upper():
                fn_norm = normalize_code(f)
                if rep_norm.replace('RFI', '') in fn_norm.replace('RFI', ''):
                    full_p = os.path.join(input_dir, f)
                    return full_p, parse_rfi_excel(full_p)['sheet_name']
            
    return None


def cleanup_headless_excel():
    """Dọn dẹp các tiến trình Excel chạy ngầm không có cửa sổ để tránh nghẽn COM hoặc khóa file."""
    try:
        import subprocess
        subprocess.run(
            ['powershell', '-NoProfile', '-Command', 
             'Get-Process excel -EA SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } | Stop-Process -Force -EA SilentlyContinue'],
            capture_output=True,
            timeout=5
        )
    except Exception:
        pass


def sanitize_and_repair_workbook(file_path: str, target_sheet: Optional[str] = None) -> str:
    """
    Tự động chuẩn hóa và sửa lỗi cấu trúc Workbook bằng openpyxl khi Excel COM bị chặn:
    - Loại bỏ các vùng ô gộp (merged_cells) bị xung đột, trùng lặp hoặc đè lên nhau (nguyên nhân gây lỗi OLE Open failed).
    - Nếu có chỉ định target_sheet và sheet đó tồn tại: chỉ giữ lại target_sheet để xuất PDF độc lập nhanh và an toàn tuyệt đối,
      loại bỏ các sheet phụ/rác bị hỏng làm treo Excel.
    - Lưu ra tệp tạm thời trong thư mục temp và trả về đường dẫn tệp đã làm sạch.
    """
    import uuid
    wb = openpyxl.load_workbook(file_path)
    
    target_match = None
    if target_sheet:
        for s in wb.sheetnames:
            if s.strip().lower() == target_sheet.strip().lower():
                target_match = s
                break
            elif target_sheet.strip().lower() in s.strip().lower() or s.strip().lower() in target_sheet.strip().lower():
                target_match = s
                break
                
    if target_match:
        for s in list(wb.sheetnames):
            if s != target_match:
                del wb[s]

    for ws in wb.worksheets:
        all_ranges = list(ws.merged_cells.ranges)
        ws.merged_cells.ranges.clear()
        claimed = set()
        for rng in all_ranges:
            conflict = False
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    if (r, c) in claimed:
                        conflict = True
                        break
                if conflict:
                    break
            if not conflict:
                for r in range(rng.min_row, rng.max_row + 1):
                    for c in range(rng.min_col, rng.max_col + 1):
                        claimed.add((r, c))
                ws.merged_cells.add(rng)

    tmp_name = f"repaired_{uuid.uuid4().hex[:8]}_{os.path.basename(file_path)}"
    tmp_path = os.path.join(tempfile.gettempdir(), tmp_name)
    wb.save(tmp_path)
    return tmp_path


def safe_open_workbook(excel_app, file_path: str, read_only: bool = True, target_sheet: Optional[str] = None) -> Tuple[object, Optional[str]]:
    """
    Mở workbook an toàn qua Excel COM.
    Tự động xử lý bỏ qua cảnh báo liên kết, gỡ bỏ Mark of the Web (Zone.Identifier),
    tự động phát hiện và khắc phục triệt để lỗi 'Open method of Workbooks class failed'
    bằng cách làm sạch các ô merged cells đè nhau hoặc cách ly sheet mục tiêu.
    Trả về Tuple (workbook_object, temp_file_path_if_repaired_else_None).
    """
    abs_path = os.path.abspath(file_path)
    
    # 1. Tự động gỡ Mark of the Web (Zone.Identifier) nếu file tải từ mạng/Zalo/Chrome
    ads = abs_path + ":Zone.Identifier"
    if os.path.exists(ads):
        try:
            os.remove(ads)
        except Exception:
            pass

    try:
        excel_app.DisplayAlerts = False
        excel_app.AskToUpdateLinks = False
    except Exception:
        pass
        
    # Thử lần 1: Mở trực tiếp với UpdateLinks=0
    try:
        wb = excel_app.Workbooks.Open(
            abs_path,
            UpdateLinks=0,
            ReadOnly=read_only,
            IgnoreReadOnlyRecommended=True
        )
        return wb, None
    except Exception:
        pass

    # Thử lần 2: CorruptLoad=1 (xlNormalLoad)
    try:
        wb = excel_app.Workbooks.Open(
            abs_path,
            UpdateLinks=0,
            ReadOnly=read_only,
            IgnoreReadOnlyRecommended=True,
            CorruptLoad=1
        )
        return wb, None
    except Exception:
        pass

    # Thử lần 3: CorruptLoad=2 (xlRepairFile - tự động sửa)
    try:
        wb = excel_app.Workbooks.Open(
            abs_path,
            UpdateLinks=0,
            ReadOnly=read_only,
            IgnoreReadOnlyRecommended=True,
            CorruptLoad=2
        )
        return wb, None
    except Exception:
        pass

    # Thử lần 4 (Tự động phục hồi & Làm sạch cấu trúc XML/Merged Cells bị lỗi bằng openpyxl)
    try:
        repaired_path = sanitize_and_repair_workbook(abs_path, target_sheet=target_sheet)
        wb = excel_app.Workbooks.Open(
            repaired_path,
            UpdateLinks=0,
            ReadOnly=read_only,
            IgnoreReadOnlyRecommended=True
        )
        return wb, repaired_path
    except Exception:
        pass

    # Thử lần 5: Mở chuẩn để hiển thị traceback gốc nếu file hoàn toàn không thể mở
    wb = excel_app.Workbooks.Open(abs_path, ReadOnly=read_only)
    return wb, None


def export_excel_sheet_to_pdf(
    excel_app,
    file_path: str,
    sheet_name: str,
    output_pdf: str,
    orientation: Optional[int] = None,
    is_rfi: bool = False
) -> None:
    """
    Xuất 1 sheet trong Excel sang PDF sử dụng Excel COM an toàn.
    Tuân thủ đúng định dạng in (PageSetup) của file Excel đưa vào:
    - Đối với Phiếu RFI: Giữ nguyên 100% định dạng và tỉ lệ in đầu vào (Zoom/Scale) của file Excel đưa vào, khổ ngang.
    - Đối với Báo cáo RP:
      + Nếu orientation is not None: Áp dụng theo định hướng chỉ định (1: Portrait, 2: Landscape).
      + Riêng biểu mẫu báo cáo RP-PPP, RP-ICR và RP-MRIR luôn luôn đảm bảo xuất theo khổ giấy ngang (Landscape = 2).
      + Thiết lập FitToPagesWide = 1 và FitToPagesTall = False: Giữ cho chiều ngang vừa vặn 1 trang (không tràn cột),
        đồng thời cho phép chiều dọc tự do hiển thị 1 trang hoặc nhiều trang theo đúng định dạng ngắt trang đã đặt trong file Excel.
      + Riêng RP-DIC: Điều chỉnh vùng in loại bỏ cột S rỗng 0.45pt để bảo toàn đường viền mép phải ô cuối.
    """
    wb, temp_repaired = safe_open_workbook(excel_app, file_path, read_only=True, target_sheet=sheet_name)
    try:
        # Tìm sheet chính xác hoặc sheet tương đồng (không phân biệt hoa/thường)
        target_ws = None
        for s in wb.Sheets:
            if s.Name.strip().lower() == sheet_name.strip().lower():
                target_ws = s
                break
        if target_ws is None:
            for s in wb.Sheets:
                if sheet_name.strip().lower() in s.Name.strip().lower() or s.Name.strip().lower() in sheet_name.strip().lower():
                    target_ws = s
                    break
        if target_ws is None:
            target_ws = wb.Sheets(1)

        fname_upper = os.path.basename(file_path).upper()
        sheet_upper = target_ws.Name.upper()
        detected_rfi = is_rfi or ('-RFI-' in fname_upper) or ('_RFI' in fname_upper) or ('RFI-' in fname_upper) or ('RFI' in sheet_upper) or fname_upper.startswith('RFI')

        try:
            target_ws.PageSetup.PaperSize = 9 # xlPaperA4
        except Exception:
            pass

        if detected_rfi:
            # Đối với Phiếu RFI: Tuân thủ đúng định dạng in và tỉ lệ in đầu vào (Zoom/Scale) của file Excel
            try:
                target_ws.PageSetup.Orientation = 2 # xlLandscape (Khổ ngang chuẩn của RFI)
            except Exception:
                pass
            # Tuyệt đối không can thiệp Zoom, FitToPages hay PrintArea của RFI để giữ nguyên tỉ lệ in gốc
        else:
            # Đối với Báo cáo RP:
            is_ppp = ('-PPP-' in fname_upper) or ('RP-PPP' in fname_upper) or ('PPP' in sheet_upper)
            is_icr = ('-ICR-' in fname_upper) or ('RP-ICR' in fname_upper) or ('ICR' in sheet_upper)
            is_mrir = ('-MRIR-' in fname_upper) or ('RP-MRIR' in fname_upper) or ('MRIR' in sheet_upper) or ('_MRIR' in fname_upper)

            if is_ppp or is_icr or is_mrir:
                # Riêng biểu mẫu RP-PPP, RP-ICR và RP-MRIR luôn luôn xuất ra theo khổ giấy ngang
                try:
                    target_ws.PageSetup.Orientation = 2 # xlLandscape
                except Exception:
                    pass
            elif orientation is not None:
                try:
                    target_ws.PageSetup.Orientation = orientation
                except Exception:
                    pass

            try:
                target_ws.PageSetup.Zoom = False
                target_ws.PageSetup.FitToPagesWide = 1
                target_ws.PageSetup.FitToPagesTall = False
            except Exception:
                pass

            # Sửa lỗi mất đường viền mép phải ô cuối cho RP-DIC (do PrintArea thừa cột S rỗng 0.45pt)
            is_rp_dic = ('RP-DIC' in fname_upper) or ('RP-DIC' in sheet_upper) or (('-DIC-' in fname_upper or '_DIC' in fname_upper or 'DIC' in sheet_upper) and not detected_rfi)
            if is_rp_dic:
                try:
                    curr_pa = str(target_ws.PageSetup.PrintArea or '')
                    if curr_pa and ('S' in curr_pa.upper()):
                        new_pa = curr_pa.replace('$S', '$R').replace(':S', ':R')
                        target_ws.PageSetup.PrintArea = new_pa
                except Exception:
                    pass
        
        target_ws.ExportAsFixedFormat(0, os.path.abspath(output_pdf))
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        if temp_repaired and os.path.exists(temp_repaired):
            try:
                os.remove(temp_repaired)
            except Exception:
                pass


def find_drawing_pdf(drawing_pdf_path: Optional[str], input_dir: str) -> Optional[str]:
    """Tìm file bản vẽ PDF thông minh."""
    if drawing_pdf_path and os.path.isfile(drawing_pdf_path):
        return os.path.abspath(drawing_pdf_path)
        
    candidates = []
    if drawing_pdf_path:
        p1 = os.path.join(input_dir, os.path.basename(drawing_pdf_path))
        if os.path.isfile(p1):
            return os.path.abspath(p1)
        parent_dir = os.path.dirname(os.path.abspath(input_dir))
        p2 = os.path.join(parent_dir, os.path.basename(drawing_pdf_path))
        if os.path.isfile(p2):
            return os.path.abspath(p2)
            
    if input_dir and os.path.isdir(input_dir):
        for f in os.listdir(input_dir):
            if f.lower().endswith('.pdf') and not f.startswith('~$'):
                candidates.append(os.path.join(input_dir, f))
            
    if input_dir:
        parent_dir = os.path.dirname(os.path.abspath(input_dir))
        if os.path.isdir(parent_dir):
            for f in os.listdir(parent_dir):
                if f.lower().endswith('.pdf') and not f.startswith('~$'):
                    candidates.append(os.path.join(parent_dir, f))
            
    for c in candidates:
        name = os.path.basename(c).lower()
        if 'ban ve' in name or 'bản vẽ' in name or 'drawing' in name or 'wmp' in name or 'ttc' in name:
            return os.path.abspath(c)
            
    if candidates:
        return os.path.abspath(candidates[0])
        
    return None


def index_drawing_sheets(drawing_pdf_path: str) -> Dict[int, str]:
    """Lập chỉ mục số Sheet của từng trang trong file bản vẽ PDF."""
    ocr = get_ocr_engine()
    doc = pymupdf.open(drawing_pdf_path)
    sheet_map = {}
    
    for pno in range(len(doc)):
        page = doc[pno]
        clip = pymupdf.Rect(550, 480, 842, 595)
        pix = page.get_pixmap(clip=clip, dpi=150)
        img = cv2.imdecode(np.frombuffer(pix.tobytes('png'), np.uint8), cv2.IMREAD_COLOR)
        res, _ = ocr(img)
        
        txt = ' '.join([item[1] for item in res]) if res else ''
        m = re.search(r'(\d{1,3})\s*/\s*(\d{1,3})', txt)
        if m:
            sheet_map[pno] = m.group(1)
        else:
            sheet_map[pno] = str(pno + 1)
            
    return sheet_map


def find_drawing_page_for_joints(doc: pymupdf.Document, joints: List[str], sheet_map: Dict[int, str]) -> int:
    """
    Tìm số thứ tự trang (0-based) trong file bản vẽ PDF phù hợp nhất cho danh sách mối hàn
    khi hồ sơ không ghi rõ số sheet (ví dụ hồ sơ chỉ có RFI).
    """
    if not joints or len(doc) <= 1:
        return 0
        
    norm_joints = [normalize_code(j) for j in joints if j]
    if not norm_joints:
        return 0
        
    # Thử 1: Tìm qua text vector
    for pno in range(len(doc)):
        txt_norm = normalize_code(doc[pno].get_text())
        if any(j in txt_norm for j in norm_joints):
            return pno

    # Thử 2: Quét OCR nhanh các trang bản vẽ
    ocr = get_ocr_engine()
    best_pno = len(doc) - 1 # Mặc định trang cuối
    best_count = 0
    
    for pno in range(len(doc)):
        page = doc[pno]
        pix = page.get_pixmap(dpi=150)
        img = cv2.imdecode(np.frombuffer(pix.tobytes('png'), np.uint8), cv2.IMREAD_COLOR)
        res, _ = ocr(img)
        if res:
            all_txt = normalize_code(' '.join([item[1] for item in res]))
            matches = sum(1 for j in norm_joints if j in all_txt)
            if matches > best_count:
                best_count = matches
                best_pno = pno
                if matches == len(norm_joints):
                    break # Khớp tất cả các mối
                    
    return best_pno


def highlight_joints_on_page(page: pymupdf.Page, joints: List[str], dpi: int = 250) -> int:
    """
    Tô màu số hiệu mối hàn trên bản vẽ với công nghệ Ink Bounding chính xác 100%,
    tuyệt đối KHÔNG lem sang các dòng chữ lân cận hoặc chi tiết khác.
    """
    ocr = get_ocr_engine()
    scale = dpi / 72.0
    pix = page.get_pixmap(dpi=dpi)
    img_gray = cv2.imdecode(np.frombuffer(pix.tobytes('png'), np.uint8), cv2.IMREAD_GRAYSCALE)
    _, thresh = cv2.threshold(img_gray, 180, 255, cv2.THRESH_BINARY_INV)
    res, _ = ocr(img_gray)
    
    if not res:
        return 0
        
    count = 0
    norm_joints = {normalize_code(j): j for j in joints}
    
    rot = page.rotation
    mb_w = page.mediabox.width
    mb_h = page.mediabox.height
    
    for item in res:
        txt = item[1]
        norm_txt = normalize_code(txt)
        box = item[0]
        
        for nj, orig_j in norm_joints.items():
            if nj in norm_txt:
                pts_x = [pt[0] for pt in box]
                pts_y = [pt[1] for pt in box]
                px0, px1 = int(min(pts_x)), int(max(pts_x))
                py0, py1 = int(min(pts_y)), int(max(pts_y))
                
                px0 = max(0, px0 - 2)
                px1 = min(img_gray.shape[1], px1 + 2)
                py0 = max(0, py0 - 2)
                py1 = min(img_gray.shape[0], py1 + 2)
                
                crop = thresh[py0:py1, px0:px1]
                
                # Chiếu dòng để xác định chính xác chiều cao thân chữ
                row_sums = np.sum(crop > 0, axis=1)
                ink_rows = np.where(row_sums > 3)[0]
                if len(ink_rows) > 0:
                    ink_py0 = py0 + np.min(ink_rows)
                    ink_py1 = py0 + np.max(ink_rows)
                else:
                    ink_py0, ink_py1 = py0, py1
                    
                # Chiếu cột để xác định chính xác chiều dài thân chữ
                crop_line = thresh[ink_py0:ink_py1+1, px0:px1]
                col_sums = np.sum(crop_line > 0, axis=0)
                ink_cols = np.where(col_sums > 0)[0]
                if len(ink_cols) > 0:
                    ink_px0 = px0 + np.min(ink_cols)
                    ink_px1 = px0 + np.max(ink_cols)
                else:
                    ink_px0, ink_px1 = px0, px1
                
                # Lề siêu nhỏ 0.2 pt ôm sát thân chữ
                margin_pt = 0.2
                vx0 = ink_px0 / scale - margin_pt
                vx1 = ink_px1 / scale + margin_pt
                vy0 = ink_py0 / scale - margin_pt
                vy1 = ink_py1 / scale + margin_pt
                
                if rot == 90:
                    ux0 = vy0
                    ux1 = vy1
                    uy0 = mb_h - vx1
                    uy1 = mb_h - vx0
                elif rot == 270:
                    ux0 = mb_w - vy1
                    ux1 = mb_w - vy0
                    uy0 = vx0
                    uy1 = vx1
                elif rot == 180:
                    ux0 = mb_w - vx1
                    ux1 = mb_w - vx0
                    uy0 = mb_h - vy1
                    uy1 = mb_h - vy0
                else:
                    ux0 = vx0
                    ux1 = vx1
                    uy0 = vy0
                    uy1 = vy1
                    
                unrot_rect = pymupdf.Rect(ux0, uy0, ux1, uy1)
                
                # Tô màu highlight sắc nét, ôm sát viền chữ, tuyệt đối không lem
                page.draw_rect(
                    unrot_rect,
                    fill=(1.0, 0.92, 0.0),
                    fill_opacity=0.45,
                    color=None,
                    overlay=True
                )
                
                count += 1
                break
                
    return count


def apply_drawing_margins_cbb(doc: pymupdf.Document, page: pymupdf.Page):
    """
    Căn chỉnh kích thước khung ngoài bản vẽ:
    - Khổ giấy in: A4 Ngang (297 x 210 mm)
    - Lề trên: 15 mm (tính từ khung ngoài bản vẽ)
    - Lề dưới: 10 mm (tính từ khung ngoài bản vẽ)
    - Lề trái: 10 mm (tính từ khung ngoài bản vẽ)
    - Lề phải: 10 mm (tính từ khung ngoài bản vẽ)
    - Chiều cao khung ngoài bản vẽ: 210 - 15 - 10 = 185 mm
    - Chiều rộng khung ngoài bản vẽ: 297 - 10 - 10 = 277 mm
    """
    page.clean_contents()
    xrefs = page.get_contents()
    if not xrefs:
        return
    xref = xrefs[0]

    # Ma trận co giãn và dịch chuyển để đạt chính xác:
    # Trên: 15.0 mm, Dưới: 10.0 mm, Trái: 10.0 mm, Phải: 10.0 mm
    Sy = 0.915009
    Sx = 0.982335
    e = 32.199
    f = -2.603

    orig_stream = doc.xref_stream(xref)
    prefix = f'q {Sy:.6f} 0 0 {Sx:.6f} {e:.3f} {f:.3f} cm\n'.encode('ascii')
    suffix = b'\nQ\n'
    new_stream = prefix + orig_stream + suffix
    doc.update_stream(xref, new_stream)


def create_reconciliation_cover(
    reconcile_data: List[Dict],
    excel_app,
    output_xlsx: str,
    output_pdf: str
):
    """
    Tạo Bảng Đối Chiếu Sự Phù Hợp (File Excel và xuất PDF trang bìa).
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Bang_Doi_Chieu'
    ws.views.sheetView[0].showGridLines = True

    # Title
    ws.merge_cells('A1:J1')
    ws['A1'] = 'DỰ ÁN ĐƯỜNG ỐNG DẪN KHÍ LÔ B - Ô MÔN (BLOCK B - O MON GAS PIPELINE PROJECT)'
    ws['A1'].font = Font(name='Arial', size=11, bold=True, color='1F497D')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells('A2:J2')
    ws['A2'] = 'BẢNG ĐỐI CHIẾU SỰ PHÙ HỢP HỒ SƠ RFI & BÁO CÁO KIỂM TRA (RP)'
    ws['A2'].font = Font(name='Arial', size=14, bold=True, color='002060')
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells('A3:J3')
    now_str = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
    matched_count = sum(1 for d in reconcile_data if d['is_match'])
    total_count = len(reconcile_data)
    ws['A3'] = f'Thời gian lập: {now_str} | Tổng số bộ hồ sơ: {total_count} | Phù hợp: {matched_count}/{total_count} | Cảnh báo lệch: {total_count - matched_count}'
    ws['A3'].font = Font(name='Arial', size=9, italic=True, color='595959')
    ws['A3'].alignment = Alignment(horizontal='center', vertical='center')

    headers = [
        'STT', 'Mã Báo cáo (RP No.)', 'Mã Thư mời (RFI No.)', 'Bản vẽ',
        'SL RP', 'Chi tiết theo RP', 'SL RFI', 'Chi tiết theo RFI',
        'Trạng thái đối chiếu', 'Ghi chú / Cảnh báo'
    ]

    thin = Side(border_style='thin', color='B0B0B0')
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    ws.row_dimensions[5].height = 26
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=5, column=c_idx, value=h)
        cell.font = Font(name='Arial', size=9, bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border

    for idx, item in enumerate(reconcile_data, 1):
        row_idx = 5 + idx
        ws.row_dimensions[row_idx].height = 24
        is_icr = item.get('is_icr', False)
        is_mrir = item.get('is_mrir', False)
        is_dic = item.get('is_dic', False)
        only_rfi = item.get('only_rfi', False)
        if is_icr:
            sheet_str = 'Không áp dụng'
            detail_rp = ', '.join(item['joints_rp']) if not only_rfi else 'Không áp dụng'
            detail_rfi = f"{item['count_rfi']} ống"
            count_rp = item['count_rp'] if not only_rfi else '-'
            count_rfi = item['count_rfi']
        elif is_mrir:
            sheet_str = 'Không áp dụng'
            detail_rp = ', '.join(item['joints_rp']) if not only_rfi else 'Không áp dụng'
            detail_rfi = 'Nghiệm thu nhận vật tư'
            count_rp = item['count_rp'] if not only_rfi else '-'
            count_rfi = item.get('count_rfi', 1)
        elif is_dic:
            sheet_str = 'Không áp dụng'
            detail_rp = ', '.join(item['joints_rp']) if not only_rfi else 'Không áp dụng'
            detail_rfi = 'Hiệu chuẩn thiết bị'
            count_rp = item['count_rp'] if not only_rfi else '-'
            count_rfi = item.get('count_rfi', 1)
        elif only_rfi:
            sheet_str = f"Sheet {item['sheet_num']}" if item.get('sheet_num') else 'Tự động'
            detail_rp = 'Không có (Chỉ RFI)'
            detail_rfi = ', '.join(item['joints_rfi']) if item['joints_rfi'] else '-'
            count_rp = '-'
            count_rfi = item['count_rfi']
        else:
            sheet_str = f"Sheet {item['sheet_num']}" if item.get('sheet_num') else 'Tự động'
            detail_rp = ', '.join(item['joints_rp'])
            detail_rfi = ', '.join(item['joints_rfi'])
            count_rp = len(item['joints_rp'])
            count_rfi = len(item['joints_rfi'])

        vals = [
            idx,
            item['report_no'],
            item['rfi_no'],
            sheet_str,
            count_rp,
            detail_rp,
            count_rfi,
            detail_rfi,
            item['status'],
            item['note']
        ]
        is_match = item['is_match']
        for col_idx, val in enumerate(vals, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name='Arial', size=9)
            cell.border = border
            if col_idx in [1, 4, 5, 7, 9]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')
                
            if col_idx == 9: # Status column
                if is_match:
                    cell.font = Font(name='Arial', size=9, bold=True, color='276A3C')
                    cell.fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
                else:
                    cell.font = Font(name='Arial', size=9, bold=True, color='C00000')
                    cell.fill = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')

    # Signatures
    sig_row = len(reconcile_data) + 8
    ws.row_dimensions[sig_row].height = 18
    ws.row_dimensions[sig_row + 1].height = 45

    ws.cell(row=sig_row, column=2, value='ĐẠI DIỆN TỔNG THẦU EPC (PTSCL18)').font = Font(name='Arial', size=9, bold=True)
    ws.cell(row=sig_row, column=2).alignment = Alignment(horizontal='center')

    ws.cell(row=sig_row, column=6, value='GIÁM SÁT THI CÔNG (IBST/S)').font = Font(name='Arial', size=9, bold=True)
    ws.cell(row=sig_row, column=6).alignment = Alignment(horizontal='center')

    ws.cell(row=sig_row, column=9, value='ĐẠI DIỆN CHỦ ĐẦU TƯ (SWPOC)').font = Font(name='Arial', size=9, bold=True)
    ws.cell(row=sig_row, column=9).alignment = Alignment(horizontal='center')

    col_widths = {1: 6, 2: 32, 3: 32, 4: 14, 5: 8, 6: 28, 7: 8, 8: 28, 9: 20, 10: 25}
    for col_idx, w in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    wb.save(output_xlsx)

    # Export to PDF via Excel COM
    wb_com, temp_cover = safe_open_workbook(excel_app, output_xlsx, read_only=True, target_sheet='Bang_Doi_Chieu')
    try:
        ws_com = wb_com.Sheets('Bang_Doi_Chieu')
        ws_com.PageSetup.PaperSize = 9 # xlPaperA4
        ws_com.PageSetup.Orientation = 2 # xlLandscape
        ws_com.PageSetup.Zoom = False
        ws_com.PageSetup.FitToPagesWide = 1
        ws_com.PageSetup.FitToPagesTall = 1
        ws_com.ExportAsFixedFormat(0, os.path.abspath(output_pdf))
    finally:
        try:
            wb_com.Close(False)
        except Exception:
            pass
        if temp_cover and os.path.exists(temp_cover):
            try:
                os.remove(temp_cover)
            except Exception:
                pass


def build_reconcile_list(input_dir: str) -> List[Dict]:
    """
    Quét thư mục INPUT và xây dựng danh sách đối chiếu hồ sơ:
    - Hồ sơ có cả RP và RFI: ghép cặp và đối chiếu mối hàn / số lượng ống.
    - Hồ sơ chỉ có RFI (không có RP): ghi nhận hồ sơ độc lập, bỏ qua RP, chỉ dùng RFI và bản vẽ (tô màu theo RFI).
    - Sắp xếp toàn bộ theo thứ tự từ nhỏ đến lớn (seq_no).
    """
    input_dir = os.path.abspath(input_dir)
    all_files = [os.path.join(input_dir, f) for f in os.listdir(input_dir) if f.lower().endswith(('.xlsx', '.xls', '.xlsm')) and not f.startswith('~$')]
    rp_files = [f for f in all_files if 'RFI' not in os.path.basename(f).upper()]
    rfi_files = [f for f in all_files if 'RFI' in os.path.basename(f).upper()]
    
    reconcile_list = []
    matched_rfi_paths = set()
    
    # 1. Ghép các file RP với RFI tương ứng
    rp_files.sort(key=lambda p: extract_seq_number(os.path.basename(p)))
    for rp_path in rp_files:
        rp_info = parse_rp_excel(rp_path)
        rfi_match = find_matching_rfi(rp_info, input_dir)
        
        rfi_path = rfi_match[0] if rfi_match else None
        rfi_sheet = rfi_match[1] if rfi_match else None
        
        if rfi_path:
            matched_rfi_paths.add(os.path.abspath(rfi_path))
            rfi_info = parse_rfi_excel(rfi_path)
            joints_rfi = rfi_info['joints']
            pipe_count_rfi = rfi_info.get('pipe_count')
        else:
            joints_rfi = []
            pipe_count_rfi = None
            
        is_icr = rp_info.get('is_icr', False)
        is_mrir = rp_info.get('is_mrir', False)
        is_dic = rp_info.get('is_dic', False)
        item_type = rp_info.get('item_type', 'PIPE' if is_icr else ('MATERIAL' if is_mrir else ('EQUIPMENT' if is_dic else 'WELD')))
        rec = reconcile_joints(rp_info['joints'], joints_rfi, item_type=item_type, rfi_pipe_count=pipe_count_rfi)
        
        item = {
            'rp_path': rp_path,
            'rp_sheet': rp_info['sheet_name'],
            'report_no': rp_info['report_no'],
            'rfi_path': rfi_path,
            'rfi_sheet': rfi_sheet,
            'rfi_no': rp_info['rfi_no'] or (rfi_info['rfi_no'] if rfi_path else 'Chưa rõ'),
            'sheet_num': rp_info['drawing_sheet_num'],
            'joints_rp': rp_info['joints'],
            'joints_rfi': joints_rfi,
            'pipe_count_rfi': pipe_count_rfi,
            'is_icr': is_icr,
            'is_mrir': is_mrir,
            'is_dic': is_dic,
            'item_type': item_type,
            'count_rp': rec['count_rp'],
            'count_rfi': rec['count_rfi'],
            'is_match': rec['is_match'],
            'status': rec['status'],
            'note': rec['note'],
            'only_rfi': False,
            'seq_no': rp_info['seq_no']
        }
        reconcile_list.append(item)
        
    # 2. Xử lý các file RFI không có file RP tương ứng (hoặc thư mục chỉ có RFI)
    unmatched_rfis = [f for f in rfi_files if os.path.abspath(f) not in matched_rfi_paths]
    unmatched_rfis.sort(key=lambda p: extract_seq_number(os.path.basename(p)))
    
    for rfi_path in unmatched_rfis:
        rfi_info = parse_rfi_excel(rfi_path)
        joints_rfi = rfi_info['joints']
        pipe_count = rfi_info.get('pipe_count')
        is_icr = ('ICR' in rfi_info['sheet_name'].upper()) or ('-ICR-' in os.path.basename(rfi_path).upper())
        is_mrir = ('MRIR' in rfi_info['sheet_name'].upper()) or ('-MRIR-' in os.path.basename(rfi_path).upper()) or ('_MRIR' in os.path.basename(rfi_path).upper())
        is_dic = ('DIC' in rfi_info['sheet_name'].upper()) or ('-DIC-' in os.path.basename(rfi_path).upper()) or ('_DIC' in os.path.basename(rfi_path).upper())
        
        if is_icr:
            item_type = 'PIPE'
            count_rfi = pipe_count if pipe_count is not None else len(joints_rfi)
            note = f"Chỉ nghiệm thu RFI ({count_rfi} ống)"
        elif is_mrir:
            item_type = 'MATERIAL'
            count_rfi = 1
            note = "Chỉ nghiệm thu RFI (Nhận vật tư)"
        elif is_dic:
            item_type = 'EQUIPMENT'
            count_rfi = 1
            note = "Chỉ nghiệm thu RFI (Hiệu chuẩn thiết bị)"
        else:
            item_type = 'WELD'
            count_rfi = len(joints_rfi)
            note = f"Chỉ nghiệm thu RFI ({count_rfi} mối hàn)"
            
        item = {
            'rp_path': None,
            'rp_sheet': None,
            'report_no': '(Không có - Chỉ RFI)',
            'rfi_path': rfi_path,
            'rfi_sheet': rfi_info['sheet_name'],
            'rfi_no': rfi_info['rfi_no'],
            'sheet_num': None, # Sẽ được đối chiếu tự động với bản vẽ nếu cần
            'joints_rp': [],
            'joints_rfi': joints_rfi,
            'pipe_count_rfi': pipe_count,
            'is_icr': is_icr,
            'is_mrir': is_mrir,
            'is_dic': is_dic,
            'item_type': item_type,
            'count_rp': 0,
            'count_rfi': count_rfi,
            'is_match': True,
            'status': 'PHÙ HỢP',
            'note': note,
            'only_rfi': True,
            'seq_no': rfi_info['seq_no']
        }
        reconcile_list.append(item)
        
    # Sắp xếp toàn bộ danh sách hồ sơ theo thứ tự từ nhỏ tới lớn
    reconcile_list.sort(key=lambda it: it.get('seq_no', 0))
    return reconcile_list


def batch_process(
    input_dir: str,
    output_dir: str,
    drawing_pdf_path: Optional[str] = None,
    output_suffix: str = "",
    progress_callback=None
) -> Dict:
    """
    Quy trình xử lý hoàn chỉnh theo yêu cầu của dự án:
    1. Sắp xếp báo cáo từ nhỏ đến lớn.
    2. Đối chiếu sự phù hợp giữa RP và RFI (số lượng & tên mối).
    3. Hỗ trợ đầy đủ trường hợp chỉ có RFI (bỏ qua RP, chỉ dùng RFI và bản vẽ đã tô màu).
    4. Tạo Trang Bìa Bảng Đối Chiếu (Excel & PDF).
    5. Xử lý từng bộ: RFI (P1) -> RP (P2, nếu có) -> Bản vẽ đã tô màu số mối (P3).
    6. Gộp nối dài thành DUY NHẤT 1 FILE PDF TỔNG theo quy tắc đặt tên dự án.
    """
    input_dir = os.path.abspath(input_dir)
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    if not HAS_WIN32COM:
        return {
            'success': False,
            'message': 'Tính năng xuất PDF gộp (RFI + RP + Bản vẽ CAD) yêu cầu chạy trực tiếp trên máy tính có cài đặt Microsoft Office Excel (Local/Cloudflare). Trên máy chủ Render (Linux), vui lòng sử dụng chế độ xuất Excel (RFI hoặc RP) hoặc truy cập qua đường dẫn nội bộ.'
        }
    
    actual_drawing_path = find_drawing_pdf(drawing_pdf_path, input_dir)

    # 1. Quét tìm và sắp xếp toàn bộ danh sách hồ sơ (RP + RFI độc lập) từ nhỏ đến lớn
    reconcile_list = build_reconcile_list(input_dir)
    if not reconcile_list:
        return {'success': False, 'message': 'Không tìm thấy file Báo cáo RP hoặc Phiếu RFI nào trong thư mục INPUT!'}

    # 2. Xác định tiền tố tên file chung
    common_prefix = ""
    for it in reconcile_list:
        rep = it['report_no']
        if rep and rep != '(Không có - Chỉ RFI)':
            m = re.match(r'^(.*-\D+)\d+$', rep)
            if m:
                common_prefix = m.group(1)
                break
        elif it.get('rfi_no'):
            rfi_code = it['rfi_no']
            m = re.match(r'^(.*-\D+)\d+$', rfi_code)
            if m:
                common_prefix = m.group(1).replace('-RFI-', '-')
                break
                
    if not common_prefix:
        common_prefix = "SWG-LLM18-70-PL-PLI-T11-CAI-"

    # Kiểm tra bản vẽ nếu có hồ sơ cần bản vẽ (Không áp dụng cho ICR, MRIR, DIC)
    has_drawing_needed = any(not (it.get('is_icr') or it.get('is_mrir') or it.get('is_dic')) for it in reconcile_list)
    if has_drawing_needed and not actual_drawing_path:
        return {'success': False, 'message': f'Không tìm thấy file bản vẽ PDF nào trong thư mục {input_dir} hoặc thư mục gốc!'}

    now = datetime.datetime.now()
    timestamp_str = now.strftime('%Y%m%d-%Hh%M')
    combined_pdf_name = f"{common_prefix}{timestamp_str}.pdf"
    combined_pdf_path = os.path.join(output_dir, combined_pdf_name)
    reconcile_xlsx_name = f"Bang_Doi_Chieu_RP_RFI_{timestamp_str}.xlsx"
    reconcile_xlsx_path = os.path.join(output_dir, reconcile_xlsx_name)

    # 3. Khởi tạo chỉ mục bản vẽ (nếu có bản vẽ)
    sheet_map = {}
    doc_drawing_source = None
    if actual_drawing_path and os.path.isfile(actual_drawing_path) and has_drawing_needed:
        if progress_callback:
            progress_callback(0, len(reconcile_list) + 2, "Đang lập chỉ mục bản vẽ...")
        sheet_map = index_drawing_sheets(actual_drawing_path)
        doc_drawing_source = pymupdf.open(actual_drawing_path)

        # Cập nhật số trang bản vẽ cho các bộ chỉ có RFI (chưa có sheet_num)
        for it in reconcile_list:
            if not (it.get('is_icr') or it.get('is_mrir') or it.get('is_dic')) and not it.get('sheet_num'):
                joints = it.get('joints_rp') or it.get('joints_rfi', [])
                target_pno = find_drawing_page_for_joints(doc_drawing_source, joints, sheet_map)
                if target_pno in sheet_map:
                    it['sheet_num'] = sheet_map[target_pno]

    # 4. Dọn dẹp tiến trình Excel headless cũ (nếu có) trước khi tạo instance mới
    cleanup_headless_excel()
    
    # 5. Khởi động Excel COM
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.AlertBeforeOverwriting = False
    excel.EnableEvents = False
    excel.ScreenUpdating = False
    
    try:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp_dir:
            # 5.1. Tạo trang bìa Bảng Đối Chiếu (PDF & Excel)
            if progress_callback:
                progress_callback(1, len(reconcile_list) + 2, "Đang tạo Bảng đối chiếu sự phù hợp RP - RFI...")
            cover_pdf_path = os.path.join(tmp_dir, 'cover_page.pdf')
            create_reconciliation_cover(reconcile_list, excel, reconcile_xlsx_path, cover_pdf_path)
            
            # 5.2. Khởi tạo Document PDF tổng nối dài
            final_doc = pymupdf.open()
            # Trang đầu tiên: Trang bìa bảng đối chiếu
            cover_doc = pymupdf.open(cover_pdf_path)
            final_doc.insert_pdf(cover_doc)
            cover_doc.close()

            # 5.3. Xử lý từng bộ theo trật tự từ nhỏ tới lớn
            for idx, item in enumerate(reconcile_list):
                disp_name = item['report_no'] if item['report_no'] != '(Không có - Chỉ RFI)' else item['rfi_no']
                if progress_callback:
                    progress_callback(idx + 2, len(reconcile_list) + 2, f"Đang xử lý bộ {idx+1}/{len(reconcile_list)}: {disp_name}...")
                    
                tmp_rfi_pdf = os.path.join(tmp_dir, f'rfi_{idx}.pdf')
                tmp_rp_pdf = os.path.join(tmp_dir, f'rp_{idx}.pdf')
                
                # A. Xuất RFI sang PDF (Theo đúng định dạng in của file Excel đưa vào, mặc định khổ ngang)
                if item['rfi_path'] and os.path.isfile(item['rfi_path']):
                    export_excel_sheet_to_pdf(excel, item['rfi_path'], item['rfi_sheet'], tmp_rfi_pdf, orientation=None, is_rfi=True)
                    doc_rfi = pymupdf.open(tmp_rfi_pdf)
                    final_doc.insert_pdf(doc_rfi)
                    doc_rfi.close()
                
                # B. Xuất RP sang PDF (chỉ xuất nếu có file RP; nếu chỉ có RFI thì bỏ qua RP)
                if item.get('rp_path') and os.path.isfile(item['rp_path']):
                    export_excel_sheet_to_pdf(excel, item['rp_path'], item['rp_sheet'], tmp_rp_pdf, orientation=None, is_rfi=False)
                    doc_rp = pymupdf.open(tmp_rp_pdf)
                    final_doc.insert_pdf(doc_rp)
                    doc_rp.close()
                
                # C. Trích xuất đúng trang bản vẽ và tô màu số mối (Không áp dụng cho hồ sơ ICR, MRIR và DIC)
                if not (item.get('is_icr') or item.get('is_mrir') or item.get('is_dic')) and doc_drawing_source:
                    sheet_num = item.get('sheet_num')
                    joints = item.get('joints_rp') if item.get('joints_rp') else item.get('joints_rfi', [])
                    target_pno = None
                    
                    if sheet_num:
                        for pno, s in sheet_map.items():
                            if str(s) == str(sheet_num):
                                target_pno = pno
                                break
                    if target_pno is None:
                        target_pno = find_drawing_page_for_joints(doc_drawing_source, joints, sheet_map)
                        
                    dw_page_doc = pymupdf.open()
                    dw_page_doc.insert_pdf(doc_drawing_source, from_page=target_pno, to_page=target_pno)
                    p = dw_page_doc[0]
                    
                    # Căn chỉnh kích thước khung ngoài bản vẽ theo quy định (Lề trên 15mm, 3 lề dưới/trái/phải 10mm)
                    apply_drawing_margins_cbb(dw_page_doc, p)

                    # Tô màu highlight chuẩn xác không lem
                    highlight_joints_on_page(p, joints, dpi=250)
                    
                    final_doc.insert_pdf(dw_page_doc)
                    dw_page_doc.close()

            if doc_drawing_source:
                doc_drawing_source.close()
                doc_drawing_source = None

            # 6. Lưu File PDF tổng nối dài duy nhất
            total_pages = len(final_doc)
            final_doc.save(combined_pdf_path)
            final_doc.close()
            
            if progress_callback:
                progress_callback(len(reconcile_list) + 2, len(reconcile_list) + 2, "Hoàn tất xuất file PDF tổng!")

    finally:
        if doc_drawing_source:
            try:
                doc_drawing_source.close()
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass
        try:
            del excel
        except Exception:
            pass
        import gc
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        
    matched_total = sum(1 for d in reconcile_list if d['is_match'])
    warn_total = len(reconcile_list) - matched_total
    
    return {
        'success': True,
        'combined_pdf': combined_pdf_path,
        'reconcile_xlsx': reconcile_xlsx_path,
        'total_sets': len(reconcile_list),
        'matched_count': matched_total,
        'warn_count': warn_total,
        'reconcile_data': reconcile_list,
        'message': f"Đã gộp thành công {len(reconcile_list)} bộ hồ sơ thành 1 File PDF nối dài: {os.path.basename(combined_pdf_path)} ({total_pages} trang)"
    }
