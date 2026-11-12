import os
import sys
import copy
import glob
import re
import datetime
import zipfile
import xml.etree.ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8')

# Đăng ký đầy đủ 38 namespace chuẩn của Word OpenXML để không bao giờ bị lỗi 'The file appears to be corrupted'
namespaces = [
    ('wpc', 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'),
    ('cx', 'http://schemas.microsoft.com/office/drawing/2014/chartex'),
    ('cx1', 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex'),
    ('cx2', 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex'),
    ('cx3', 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex'),
    ('cx4', 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex'),
    ('cx5', 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex'),
    ('cx6', 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex'),
    ('cx7', 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex'),
    ('cx8', 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex'),
    ('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006'),
    ('aink', 'http://schemas.microsoft.com/office/drawing/2016/ink'),
    ('am3d', 'http://schemas.microsoft.com/office/drawing/2017/model3d'),
    ('o', 'urn:schemas-microsoft-com:office:office'),
    ('oel', 'http://schemas.microsoft.com/office/2019/extlst'),
    ('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
    ('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math'),
    ('v', 'urn:schemas-microsoft-com:vml'),
    ('wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'),
    ('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'),
    ('w10', 'urn:schemas-microsoft-com:office:word'),
    ('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'),
    ('w14', 'http://schemas.microsoft.com/office/word/2010/wordml'),
    ('w15', 'http://schemas.microsoft.com/office/word/2012/wordml'),
    ('w16cex', 'http://schemas.microsoft.com/office/word/2018/wordml/cex'),
    ('w16cid', 'http://schemas.microsoft.com/office/word/2016/wordml/cid'),
    ('w16', 'http://schemas.microsoft.com/office/word/2018/wordml'),
    ('w16du', 'http://schemas.microsoft.com/office/word/2023/wordml/word16du'),
    ('w16sdtdh', 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash'),
    ('w16sdtfl', 'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock'),
    ('w16se', 'http://schemas.microsoft.com/office/word/2015/wordml/symex'),
    ('wpg', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'),
    ('wpi', 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk'),
    ('wne', 'http://schemas.microsoft.com/office/word/2006/wordml'),
    ('wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'),
    ('a', 'http://schemas.openxmlformats.org/drawingml/2006/main'),
    ('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture'),
    ('a14', 'http://schemas.microsoft.com/office/drawing/2010/main')
]

for p, u in namespaces:
    ET.register_namespace(p, u)

w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
w14_ns = 'http://schemas.microsoft.com/office/word/2010/wordml'
r_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ns = {'w': w_ns, 'w14': w14_ns, 'r': r_ns}

# XML Footer chuẩn theo mẫu đính kèm (có đường kẻ ngang, mã biểu mẫu bên trái, số trang tự động bên phải)
FOOTER_XML_TEMPLATE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Footer"/>
      <w:pBdr>
        <w:top w:val="single" w:sz="6" w:space="4" w:color="auto"/>
      </w:pBdr>
      <w:tabs>
        <w:tab w:val="right" w:pos="10206"/>
      </w:tabs>
      <w:spacing w:before="60" w:after="0"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="18"/>
        <w:szCs w:val="18"/>
      </w:rPr>
      <w:t>Biểu mẫu số/</w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:i/>
        <w:sz w:val="18"/>
        <w:szCs w:val="18"/>
      </w:rPr>
      <w:t>Form No:</w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="18"/>
        <w:szCs w:val="18"/>
      </w:rPr>
      <w:t xml:space="preserve"> SWG-AAAA-QC-BBB-GF-0020</w:t>
    </w:r>
    <w:r>
      <w:tab/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="18"/>
        <w:szCs w:val="18"/>
      </w:rPr>
      <w:t>Pages:</w:t>
    </w:r>
    <w:fldSimple w:instr="PAGE">
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
          <w:sz w:val="18"/>
          <w:szCs w:val="18"/>
        </w:rPr>
        <w:t>1</w:t>
      </w:r>
    </w:fldSimple>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="18"/>
        <w:szCs w:val="18"/>
      </w:rPr>
      <w:t>/</w:t>
    </w:r>
    <w:fldSimple w:instr="NUMPAGES">
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
          <w:sz w:val="18"/>
          <w:szCs w:val="18"/>
        </w:rPr>
        <w:t>1</w:t>
      </w:r>
    </w:fldSimple>
  </w:p>
</w:ftr>'''

# RapidOCR optional import for maximum portability
try:
    from rapidocr_onnxruntime import RapidOCR
    HAS_OCR = True
    ocr_engine = None
except Exception:
    HAS_OCR = False
    ocr_engine = None

def get_ocr():
    global ocr_engine
    if HAS_OCR and ocr_engine is None:
        try:
            ocr_engine = RapidOCR()
        except Exception as e:
            print(f"Lưu ý: Không khởi tạo được RapidOCR: {e}")
    return ocr_engine

# Helper function to set cell text and formatting
def set_cell_text(tc, text, align='center', bold=False):
    p = tc.find('w:p', ns)
    if p is None:
        p = ET.SubElement(tc, f'{{{w_ns}}}p')
    
    pPr = p.find('w:pPr', ns)
    if pPr is None:
        pPr = ET.SubElement(p, f'{{{w_ns}}}pPr')
    
    jc = pPr.find('w:jc', ns)
    if jc is None:
        jc = ET.SubElement(pPr, f'{{{w_ns}}}jc')
    jc.attrib[f'{{{w_ns}}}val'] = align
    
    runs = p.findall('w:r', ns)
    first_rPr = None
    if runs and runs[0].find('w:rPr', ns) is not None:
        first_rPr = copy.deepcopy(runs[0].find('w:rPr', ns))
    
    for r in runs:
        p.remove(r)
        
    if text:
        r = ET.SubElement(p, f'{{{w_ns}}}r')
        if first_rPr is not None:
            r.append(first_rPr)
            if bold:
                if r.find('.//w:b', ns) is None:
                    ET.SubElement(r.find('w:rPr', ns), f'{{{w_ns}}}b')
                if r.find('.//w:bCs', ns) is None:
                    ET.SubElement(r.find('w:rPr', ns), f'{{{w_ns}}}bCs')
        else:
            rPr = ET.SubElement(r, f'{{{w_ns}}}rPr')
            rFonts = ET.SubElement(rPr, f'{{{w_ns}}}rFonts')
            rFonts.attrib[f'{{{w_ns}}}ascii'] = 'Times New Roman'
            rFonts.attrib[f'{{{w_ns}}}hAnsi'] = 'Times New Roman'
            rFonts.attrib[f'{{{w_ns}}}cs'] = 'Times New Roman'
            sz = ET.SubElement(rPr, f'{{{w_ns}}}sz')
            sz.attrib[f'{{{w_ns}}}val'] = '20'
            szCs = ET.SubElement(rPr, f'{{{w_ns}}}szCs')
            szCs.attrib[f'{{{w_ns}}}val'] = '20'
            lang = ET.SubElement(rPr, f'{{{w_ns}}}lang')
            lang.attrib[f'{{{w_ns}}}val'] = 'vi-VN'
            if bold:
                ET.SubElement(rPr, f'{{{w_ns}}}b')
                ET.SubElement(rPr, f'{{{w_ns}}}bCs')
                
        t = ET.SubElement(r, f'{{{w_ns}}}t')
        t.text = str(text)
        if str(text).startswith(' ') or str(text).endswith(' '):
            t.attrib['{http://www.w3.org/XML/1998/namespace}space'] = 'preserve'

# Helper function to set Section 7 plan items with bullet list style
def set_sec7_cell(tc, text):
    p = tc.find('w:p', ns)
    if p is None:
        p = ET.SubElement(tc, f'{{{w_ns}}}p')
        
    pPr = p.find('w:pPr', ns)
    if pPr is None:
        pPr = ET.SubElement(p, f'{{{w_ns}}}pPr')
        
    runs = p.findall('w:r', ns)
    for r in runs:
        p.remove(r)
        
    if not text:
        pStyle = pPr.find('w:pStyle', ns)
        if pStyle is not None: pPr.remove(pStyle)
        numPr = pPr.find('w:numPr', ns)
        if numPr is not None: pPr.remove(numPr)
        return

    pStyle = pPr.find('w:pStyle', ns)
    if pStyle is None:
        pStyle = ET.SubElement(pPr, f'{{{w_ns}}}pStyle')
        pStyle.attrib[f'{{{w_ns}}}val'] = 'ListParagraph'
    numPr = pPr.find('w:numPr', ns)
    if numPr is None:
        numPr = ET.SubElement(pPr, f'{{{w_ns}}}numPr')
        ilvl = ET.SubElement(numPr, f'{{{w_ns}}}ilvl')
        ilvl.attrib[f'{{{w_ns}}}val'] = '0'
        numId = ET.SubElement(numPr, f'{{{w_ns}}}numId')
        numId.attrib[f'{{{w_ns}}}val'] = '17'
        
    r = ET.SubElement(p, f'{{{w_ns}}}r')
    rPr = ET.SubElement(r, f'{{{w_ns}}}rPr')
    rFonts = ET.SubElement(rPr, f'{{{w_ns}}}rFonts')
    rFonts.attrib[f'{{{w_ns}}}ascii'] = 'Times New Roman'
    rFonts.attrib[f'{{{w_ns}}}hAnsi'] = 'Times New Roman'
    rFonts.attrib[f'{{{w_ns}}}cs'] = 'Times New Roman'
    sz = ET.SubElement(rPr, f'{{{w_ns}}}sz')
    sz.attrib[f'{{{w_ns}}}val'] = '20'
    szCs = ET.SubElement(rPr, f'{{{w_ns}}}szCs')
    szCs.attrib[f'{{{w_ns}}}val'] = '20'
    lang = ET.SubElement(rPr, f'{{{w_ns}}}lang')
    lang.attrib[f'{{{w_ns}}}val'] = 'vi-VN'
    
    t = ET.SubElement(r, f'{{{w_ns}}}t')
    t.text = str(text)

def build_tblA(template, day_data):
    tbl = copy.deepcopy(template)
    trs = tbl.findall('w:tr', ns)
    
    # 1. Update Date in Row 0 Cell 1
    r0c1 = trs[0].findall('w:tc', ns)[1]
    ps = r0c1.findall('w:p', ns)
    if len(ps) >= 2:
        date_p = ps[1]
        for r in date_p.findall('w:r', ns):
            date_p.remove(r)
        r = ET.SubElement(date_p, f'{{{w_ns}}}r')
        rPr = ET.SubElement(r, f'{{{w_ns}}}rPr')
        rFonts = ET.SubElement(rPr, f'{{{w_ns}}}rFonts')
        rFonts.attrib[f'{{{w_ns}}}ascii'] = 'Times New Roman'
        rFonts.attrib[f'{{{w_ns}}}eastAsia'] = 'Times New Roman'
        rFonts.attrib[f'{{{w_ns}}}hAnsi'] = 'Times New Roman'
        rFonts.attrib[f'{{{w_ns}}}cs'] = 'Times New Roman'
        ET.SubElement(rPr, f'{{{w_ns}}}i')
        sz = ET.SubElement(rPr, f'{{{w_ns}}}sz')
        sz.attrib[f'{{{w_ns}}}val'] = '20'
        szCs = ET.SubElement(rPr, f'{{{w_ns}}}szCs')
        szCs.attrib[f'{{{w_ns}}}val'] = '20'
        lang = ET.SubElement(rPr, f'{{{w_ns}}}lang')
        lang.attrib[f'{{{w_ns}}}val'] = 'vi-VN'
        
        t = ET.SubElement(r, f'{{{w_ns}}}t')
        t.text = f"Ngày/Date:{day_data['date']}"
        
    # 2. Update Weather checkboxes in Row 7
    r7 = trs[7]
    syms = r7.findall('.//w:sym', ns)
    weather = day_data.get('weather')
    if len(syms) >= 3:
        if weather == 'Sunny':
            syms[0].attrib[f'{{{w_ns}}}font'] = 'Symbol'
            syms[0].attrib[f'{{{w_ns}}}char'] = 'F07F'
            syms[1].attrib[f'{{{w_ns}}}font'] = 'Wingdings'
            syms[1].attrib[f'{{{w_ns}}}char'] = 'F0FE'
            syms[2].attrib[f'{{{w_ns}}}font'] = 'Symbol'
            syms[2].attrib[f'{{{w_ns}}}char'] = 'F07F'
        elif weather == 'Raining':
            syms[0].attrib[f'{{{w_ns}}}font'] = 'Symbol'
            syms[0].attrib[f'{{{w_ns}}}char'] = 'F07F'
            syms[1].attrib[f'{{{w_ns}}}font'] = 'Symbol'
            syms[1].attrib[f'{{{w_ns}}}char'] = 'F07F'
            syms[2].attrib[f'{{{w_ns}}}font'] = 'Wingdings'
            syms[2].attrib[f'{{{w_ns}}}char'] = 'F0FE'
        else: # Normal / Bình thường
            syms[0].attrib[f'{{{w_ns}}}font'] = 'Wingdings'
            syms[0].attrib[f'{{{w_ns}}}char'] = 'F0FE'
            syms[1].attrib[f'{{{w_ns}}}font'] = 'Symbol'
            syms[1].attrib[f'{{{w_ns}}}char'] = 'F07F'
            syms[2].attrib[f'{{{w_ns}}}font'] = 'Symbol'
            syms[2].attrib[f'{{{w_ns}}}char'] = 'F07F'

    # 3. Update Manpower in Row 13 (Worker = Tổng cộng)
    r13 = trs[13]
    cells13 = r13.findall('w:tc', ns)
    worker_val = str(day_data['worker']) if day_data.get('worker') is not None else ""
    set_cell_text(cells13[2], worker_val, align='center')
    set_cell_text(cells13[4], worker_val, align='center')

    return tbl

def build_tblB(template, day_data):
    tbl = copy.deepcopy(template)
    trs = tbl.findall('w:tr', ns)
    sec4_items = day_data.get('sec4', [])
    
    # 9 data rows: trs[2] to trs[10]
    for idx in range(9):
        tr = trs[2 + idx]
        cells = tr.findall('w:tc', ns)
        if idx < len(sec4_items):
            item = sec4_items[idx]
            set_cell_text(cells[0], str(item.get('stt', '')), align='center')
            set_cell_text(cells[1], item.get('content', ''), align='left')
            set_cell_text(cells[2], item.get('unit', ''), align='center')
            qty_str = str(item.get('qty', '')) if item.get('qty') is not None else ''
            accum_str = str(item.get('accum', '')) if item.get('accum') is not None else ''
            set_cell_text(cells[3], qty_str, align='center')
            set_cell_text(cells[4], accum_str, align='center')
        else:
            set_cell_text(cells[0], '', align='center')
            set_cell_text(cells[1], '', align='left')
            set_cell_text(cells[2], '', align='center')
            set_cell_text(cells[3], '', align='center')
            set_cell_text(cells[4], '', align='center')
            
    return tbl

def build_tblC(template, day_data):
    tbl = copy.deepcopy(template)
    trs = tbl.findall('w:tr', ns)
    sec7_items = day_data.get('sec7', [])
    
    # 6 plan rows: trs[10] to trs[15]
    for idx in range(6):
        tr = trs[10 + idx]
        cell = tr.find('w:tc', ns)
        if idx < len(sec7_items):
            set_sec7_cell(cell, sec7_items[idx])
        else:
            set_sec7_cell(cell, '')
            
    return tbl

def find_file(pattern, default_name):
    matches = glob.glob(pattern)
    if matches:
        for m in matches:
            if os.path.basename(m).lower() == default_name.lower():
                return m
        return matches[0]
    return default_name

def clean_report_code(text):
    """Xóa bỏ các mã số báo cáo trong ngoặc đơn như (MRIR-...), (ICR-...), (FUR-...)"""
    if not text:
        return ""
    # Xóa pattern dạng (XXX-YYYY-...) hoặc (XXX) hoặc (quanity : ...)
    cleaned = re.sub(r'\s*\([^)]*-[^)]*\)', '', text)
    cleaned = re.sub(r'\s*\((?:MRIR|ICR|FUR|VSR|NDT|DIC|SPI|CAI|FTI|HTI|MTR|PPP|MEC)[^)]*\)', '', cleaned, flags=re.I)
    cleaned = re.sub(r'\s*\(\s*quanity\s*:[^)]*\)', '', cleaned, flags=re.I)
    return cleaned.strip()

def scan_zalo_reports(zalo_dir, target_dates=None):
    """Quét thư mục ảnh báo cáo Zalo để trích xuất thời tiết và nhân công bãi TTC11, ưu tiên lọc theo dải ngày thực tế"""
    reports = {}
    if not zalo_dir or not os.path.isdir(zalo_dir):
        return reports
        
    ocr = get_ocr()
    if not ocr:
        print("  [OCR] Không có thư viện OCR, bỏ qua quét ảnh Zalo.", flush=True)
        return reports
        
    print(f"  [OCR] Đang quét ảnh báo cáo tại: {zalo_dir}...", flush=True)
    valid_exts = ('.jpg', '.jpeg', '.png', '.bmp')
    all_files = [f for f in os.listdir(zalo_dir) if f.lower().endswith(valid_exts)]
    
    # Lọc ảnh theo khoảng ngày nếu có target_dates
    target_date_objs = []
    if target_dates:
        for td in target_dates:
            try:
                target_date_objs.append(datetime.datetime.strptime(td, '%d/%m/%Y').date())
            except Exception:
                pass
                
    if target_date_objs:
        min_d = min(target_date_objs) - datetime.timedelta(days=1)
        max_d = max(target_date_objs) + datetime.timedelta(days=1)
        filtered_files = []
        for f in all_files:
            p = os.path.join(zalo_dir, f)
            mt = os.path.getmtime(p)
            dt = datetime.datetime.fromtimestamp(mt).date()
            if min_d <= dt <= max_d:
                filtered_files.append((mt, f))
        filtered_files.sort(key=lambda x: x[0], reverse=True)
        image_files = [x[1] for x in filtered_files]
        print(f"  [OCR] Đã lọc {len(image_files)} ảnh Zalo trong dải ngày {min_d} đến {max_d}", flush=True)
    else:
        all_files.sort(key=lambda x: os.path.getmtime(os.path.join(zalo_dir, x)), reverse=True)
        image_files = all_files[:120]
    
    for img_name in image_files:
        img_path = os.path.join(zalo_dir, img_name)
        try:
            res, _ = ocr(img_path)
            if not res:
                continue
            texts = [r[1] for r in res]
            full_txt = ' '.join(texts).upper()
            
            # Chỉ lấy ảnh có tiêu đề TTC11
            if 'TTC11' not in full_txt or ('BÁO CÁO' not in full_txt and 'BAO CAO' not in full_txt):
                continue
                
            # Trích xuất Ngày
            date_found = None
            date_match = re.search(r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})', full_txt)
            if date_match:
                d, m, y = date_match.groups()
                date_found = f"{int(d):02d}/{int(m):02d}/{y}"
                
            if not date_found or date_found in reports:
                continue
                
            # Trích xuất Thời tiết
            weather = 'Normal'
            if 'NẮNG' in full_txt or 'NANG' in full_txt:
                weather = 'Sunny'
            elif 'MƯA' in full_txt or 'MUA' in full_txt:
                weather = 'Raining'
                
            # Trích xuất Tổng cộng nhân công
            worker_total = None
            for idx, line in enumerate(texts):
                line_clean = line.strip().upper()
                if 'TỔNG CỘNG' in line_clean or 'TONG CONG' in line_clean:
                    # Kiểm tra dòng tiếp theo (thường chứa số tổng cộng)
                    if idx + 1 < len(texts) and texts[idx+1].strip().isdigit():
                        worker_total = int(texts[idx+1].strip())
                        break
                    # Hoặc số nằm phía sau chữ TỔNG CỘNG trên cùng dòng
                    parts = re.split(r'TỔNG\s*CỘNG|TONG\s*CONG', line_clean)
                    if len(parts) > 1:
                        right_nums = re.findall(r'\b\d{1,2}\b', parts[1])
                        if right_nums:
                            worker_total = int(right_nums[-1])
                            break
                    # Fallback tìm số khác STT
                    nums = re.findall(r'\b\d{1,2}\b', line)
                    if nums:
                        worker_total = int(nums[-1])
                    break
                    
            if date_found:
                # Công nhân/Worker = Tổng cộng Zalo - 3 (trừ 3 vị trí gián tiếp: Giám sát, QC, An toàn)
                worker_val = (worker_total - 3) if worker_total is not None else None
                reports[date_found] = {
                    'weather': weather,
                    'worker': worker_val,
                    'file': img_name
                }
                print(f"  -> Tìm thấy Báo cáo {date_found}: Thời tiết={weather}, Tổng={worker_total} -> Công nhân (Tổng-3)={worker_val} ({img_name})", flush=True)
        except Exception as e:
            continue
            
    return reports

# Cơ sở dữ liệu chuẩn đã xác nhận cho Tháng 10 và Tháng 11 (100% theo FORM CHOOSE REPORT)
KNOWN_DAYS_DATA = {
    '22/10/2026': {'date': '22/10/2026', 'weather': 'Normal', 'worker': None, 'sec4': [{'stt': 1, 'content': 'Tiếp nhận vật tư ống', 'unit': 'Pcs', 'qty': 32, 'accum': 32}, {'stt': 2, 'content': 'Kiểm tra ống đầu vào và vệ sinh làm sạch lòng ống', 'unit': 'Pcs', 'qty': 16, 'accum': 16}, {'stt': 3, 'content': 'Nghiệm thu Fit-up mối hàn: KP39-W45', 'unit': 'Joint', 'qty': 1, 'accum': 1}, {'stt': 4, 'content': 'Hàn mối hàn: KP39-W45', 'unit': 'Joint', 'qty': 1, 'accum': 1}, {'stt': 5, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}], 'sec7': ['Nhận vật tư ống', 'Vệ sinh ống', 'Fit-up và hàn mối hàn ống', 'Sắp xếp dọn dẹp vệ sinh bãi']},
    '23/10/2026': {'date': '23/10/2026', 'weather': 'Sunny', 'worker': 30, 'sec4': [{'stt': 1, 'content': 'Kiểm tra ống đầu vào và vệ sinh làm sạch lòng ống', 'unit': 'Pcs', 'qty': 16, 'accum': 32}, {'stt': 2, 'content': 'Nghiệm thu tiếp nhận vật liệu hàn', 'unit': 'Lot', 'qty': 1, 'accum': 1}, {'stt': 3, 'content': 'Nghiệm thu Fit-up mối hàn: KP39-W43; KP39-W44', 'unit': 'Joint', 'qty': 2, 'accum': 3}, {'stt': 4, 'content': 'Hàn mối hàn: KP39-W43; KP39-W44', 'unit': 'Joint', 'qty': 2, 'accum': 3}, {'stt': 5, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}], 'sec7': ['Nhận vật tư ống', 'Vệ sinh ống', 'Fit-up và hàn các vị trí nối ống', 'Nghiệm thu Visual, NDT mối hàn', 'Sắp xếp dọn dẹp vệ sinh bãi']},
    '24/10/2026': {'date': '24/10/2026', 'weather': 'Sunny', 'worker': 30, 'sec4': [{'stt': 1, 'content': 'Nghiệm thu Fit-up mối hàn: KP39-W35; KP39-W37; KP39-W41; KP39-W42', 'unit': 'Joint', 'qty': 4, 'accum': 7}, {'stt': 2, 'content': 'Hàn mối hàn: KP39-W45; KP39-W44; KP39-W43', 'unit': 'Joint', 'qty': 3, 'accum': 3}, {'stt': 3, 'content': 'Nghiệm thu Visual mối hàn: KP39-W45; KP39-W44; KP39-W43', 'unit': 'Joint', 'qty': 3, 'accum': 3}, {'stt': 4, 'content': 'Nghiệm thu NDT AUT mối hàn: KP39-W45; KP39-W44; KP39-W43', 'unit': 'Joint', 'qty': 3, 'accum': 3}, {'stt': 5, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}], 'sec7': ['Nhận vật tư ống', 'Vệ sinh ống', 'Fit-up và hàn các vị trí nối ống', 'Dán FJC ống', 'Sắp xếp dọn dẹp vệ sinh bãi']},
    '25/10/2026': {'date': '25/10/2026', 'weather': 'Normal', 'worker': None, 'sec4': [{'stt': 1, 'content': 'Kiểm tra thiết bị thi công', 'unit': 'Bộ', 'qty': 1, 'accum': 1}, {'stt': 2, 'content': 'Nghiệm thu Fit-up & hàn mối hàn: KP39-W40; KP39-W39; KP39-W31; KP39-W23', 'unit': 'Joint', 'qty': 4, 'accum': 11}, {'stt': 3, 'content': 'Hàn mối hàn: KP39-W40; KP39-W39; KP39-W31; KP39-W23', 'unit': 'Joint', 'qty': 4, 'accum': 7}, {'stt': 4, 'content': 'Nghiệm thu Visual & NDT AUT: KP39-W35; KP39-W37; KP39-W41; KP39-W42', 'unit': 'Joint', 'qty': 4, 'accum': 7}, {'stt': 5, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': 1}, {'stt': 6, 'content': 'Làm sạch bề mặt, dán FJC, thử nghiệm: KP39-W45; KP39-W44; KP39-W43', 'unit': 'Joint', 'qty': 3, 'accum': 3}, {'stt': 7, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}], 'sec7': ['Vệ sinh ống', 'Fit-up và hàn ống', 'Dán FJC ống', 'Phun Foam bọc cách nhiệt', 'Đẩy ống xuống hào', 'Sắp xếp dọn dẹp vệ sinh bãi']},
    '26/10/2026': {'date': '26/10/2026', 'weather': 'Sunny', 'worker': 29, 'sec4': [{'stt': 1, 'content': 'Kiểm tra ống đầu vào và vệ sinh làm sạch lòng ống', 'unit': 'Pcs', 'qty': 8, 'accum': 40}, {'stt': 2, 'content': 'Nghiệm thu Fit-up & hàn mối hàn: KP39-W25; KP39-W27; KP39-W29; KP39-W33', 'unit': 'Joint', 'qty': 4, 'accum': 15}, {'stt': 3, 'content': 'Nghiệm thu Visual & NDT AUT mối hàn: KP39-W23; KP39-W31; KP39-W39; KP39-W40', 'unit': 'Joint', 'qty': 4, 'accum': 11}, {'stt': 4, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': 2}, {'stt': 5, 'content': 'Làm sạch bề mặt, dán FJC, thử nghiệm: KP39-W40; KP39-W41; KP39-W42', 'unit': 'Joint', 'qty': 3, 'accum': 6}, {'stt': 6, 'content': 'Nghiệm thu bọc cách nhiệt HDPUF: KP39-W43; KP39-W44; KP39-W45', 'unit': 'Joint', 'qty': 3, 'accum': 3}, {'stt': 7, 'content': 'Nghiệm thu kéo đẩy ống xuống hào: Ống 0414; 5058; 3433; 2971', 'unit': 'Ống', 'qty': 4, 'accum': 4}, {'stt': 8, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}], 'sec7': ['Vệ sinh ống', 'Fit-up và hàn ống', 'Dán FJC ống', 'Phun Foam ống', 'Kéo đẩy ống xuống hào', 'Sắp xếp dọn dẹp vệ sinh bãi']},
    '27/10/2026': {'date': '27/10/2026', 'weather': 'Sunny', 'worker': 28, 'sec4': [{'stt': 1, 'content': 'Kiểm tra ống đầu vào và vệ sinh làm sạch lòng ống', 'unit': 'Pcs', 'qty': 8, 'accum': 48}, {'stt': 2, 'content': 'Nghiệm thu Fit-up & hàn mối hàn: KP39-W38; KP39-W36; KP39-W21', 'unit': 'Joint', 'qty': 3, 'accum': 18}, {'stt': 3, 'content': 'Nghiệm thu Visual & NDT AUT mối hàn: KP39-W25; KP39-W27; KP39-W29; KP39-W33; KP39-W38; KP39-W36', 'unit': 'Joint', 'qty': 6, 'accum': 17}, {'stt': 4, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': 3}, {'stt': 5, 'content': 'Làm sạch bề mặt, dán FJC, thử nghiệm: KP39-W39; KP39-W23', 'unit': 'Joint', 'qty': 2, 'accum': 8}, {'stt': 6, 'content': 'Nghiệm thu bọc cách nhiệt HDPUF: KP39-W39; KP39-W40; KP39-W41; KP39-W42', 'unit': 'Joint', 'qty': 4, 'accum': 7}, {'stt': 7, 'content': 'Nghiệm thu kéo đẩy ống xuống hào: Ống 2971; 6991; 7254; 4187', 'unit': 'Ống', 'qty': 4, 'accum': 8}, {'stt': 8, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}], 'sec7': ['Nhận vật tư ống', 'Vệ sinh ống', 'Fit-up và hàn ống', 'Dán FJC', 'Phun foam', 'Sắp xếp dọn dẹp vệ sinh bãi']},
    '28/10/2026': {'date': '28/10/2026', 'weather': 'Sunny', 'worker': 29, 'sec4': [{'stt': 1, 'content': 'Tiếp nhận vật tư ống', 'unit': 'Pcs', 'qty': 32, 'accum': 64}, {'stt': 2, 'content': 'Kiểm tra ống đầu vào và vệ sinh làm sạch lòng ống', 'unit': 'Pcs', 'qty': 8, 'accum': 56}, {'stt': 3, 'content': 'Nghiệm thu Fit-up & hàn mối hàn: KP39-W17; KP39-W19; KP39-W32; KP39-W34', 'unit': 'Joint', 'qty': 4, 'accum': 22}, {'stt': 4, 'content': 'Nghiệm thu Visual & NDT AUT mối hàn: KP39-W32; KP39-W34; KP39-W21', 'unit': 'Joint', 'qty': 3, 'accum': 20}, {'stt': 5, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': 4}, {'stt': 6, 'content': 'Làm sạch bề mặt, dán FJC, thử nghiệm: KP39-W29; KP39-W35; KP39-W36; KP39-W37; KP39-W38', 'unit': 'Joint', 'qty': 5, 'accum': 13}, {'stt': 7, 'content': 'Nghiệm thu bọc cách nhiệt HDPUF: KP39-W35; KP39-W36; KP39-W37; KP39-W38', 'unit': 'Joint', 'qty': 4, 'accum': 11}, {'stt': 8, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}], 'sec7': ['Vệ sinh ống', 'Fit-up và hàn ống', 'Dán FJC ống', 'Phun Foam ống', 'Kéo đẩy ống xuống hào', 'Sắp xếp dọn dẹp vệ sinh bãi']},
    '29/10/2026': {'date': '29/10/2026', 'weather': 'Sunny', 'worker': 35, 'sec4': [{'stt': 1, 'content': 'Nghiệm thu Fit-up & hàn mối hàn: KP39-W11; KP39-W13; KP39-W15; KP39-W30', 'unit': 'Joint', 'qty': 4, 'accum': 26}, {'stt': 2, 'content': 'Nghiệm thu Visual & NDT AUT mối hàn: KP39-W17; KP39-W19; KP39-W30', 'unit': 'Joint', 'qty': 3, 'accum': 23}, {'stt': 3, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': 5}, {'stt': 4, 'content': 'Làm sạch bề mặt, dán FJC, thử nghiệm: KP39-W25; KP39-W27; KP39-W31; KP39-W33', 'unit': 'Joint', 'qty': 4, 'accum': 17}, {'stt': 5, 'content': 'Nghiệm thu bọc cách nhiệt HDPUF: KP39-W29; KP39-W31; KP39-W33', 'unit': 'Joint', 'qty': 3, 'accum': 14}, {'stt': 6, 'content': 'Nghiệm thu kéo đẩy ống xuống hào: Ống 4187; 7173; 7639; 3542', 'unit': 'Ống', 'qty': 4, 'accum': 12}, {'stt': 7, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}], 'sec7': ['Vệ sinh ống', 'Fit-up và hàn ống', 'Dán FJC ống', 'Phun foam', 'Kéo đẩy ống xuống hào', 'Sắp xếp dọn dẹp vệ sinh bãi']},
    '30/10/2026': {'date': '30/10/2026', 'weather': 'Raining', 'worker': 32, 'sec4': [{'stt': 1, 'content': 'Kiểm tra ống đầu vào và vệ sinh làm sạch lòng ống', 'unit': 'Pcs', 'qty': 8, 'accum': 64}, {'stt': 2, 'content': 'Nghiệm thu Fit-up & hàn mối hàn: KP39-W09; KP39-W28; KP39-W30', 'unit': 'Joint', 'qty': 3, 'accum': 29}, {'stt': 3, 'content': 'Nghiệm thu Visual & NDT AUT mối hàn: KP39-W11; KP39-W13; KP39-W15', 'unit': 'Joint', 'qty': 3, 'accum': 26}, {'stt': 4, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': 6}, {'stt': 5, 'content': 'Làm sạch bề mặt, dán FJC, thử nghiệm: KP39-W19; KP39-W21; KP39-W27', 'unit': 'Joint', 'qty': 3, 'accum': 20}, {'stt': 6, 'content': 'Nghiệm thu bọc cách nhiệt HDPUF: KP39-W29', 'unit': 'Joint', 'qty': 1, 'accum': 15}, {'stt': 7, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}], 'sec7': ['Kéo đẩy chuỗi ống xuống hào', 'Tiếp nhận và vệ sinh ống', 'Fit-up và hàn ống', 'NDT mối hàn', 'Dán FJC ống', 'Sắp xếp dọn dẹp bãi']},
    '31/10/2026': {'date': '31/10/2026', 'weather': 'Sunny', 'worker': 34, 'sec4': [{'stt': 1, 'content': 'Kiểm tra ống đầu vào và vệ sinh làm sạch lòng ống', 'unit': 'Pcs', 'qty': 8, 'accum': 72}, {'stt': 2, 'content': 'Nghiệm thu kéo đẩy ống xuống hào: Ống 7639; 3568; 0613; 3583; 3420; 4957; 7253', 'unit': 'Ống', 'qty': 7, 'accum': 19}, {'stt': 3, 'content': 'Nghiệm thu Fit-up & hàn mối hàn: KP39-W05; KP39-W07; KP39-W24; KP39-W26', 'unit': 'Joint', 'qty': 4, 'accum': 33}, {'stt': 4, 'content': 'Nghiệm thu Visual mối hàn: KP39-W09; KP39-W24; KP39-W26', 'unit': 'Joint', 'qty': 3, 'accum': 30}, {'stt': 5, 'content': 'Nghiệm thu NDT AUT mối hàn: KP39-W28; KP39-W30', 'unit': 'Joint', 'qty': 2, 'accum': 27}, {'stt': 6, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': 7}, {'stt': 7, 'content': 'Làm sạch bề mặt, dán FJC, thử nghiệm: KP39-W28; KP39-W30', 'unit': 'Joint', 'qty': 2, 'accum': 22}, {'stt': 8, 'content': 'Nghiệm thu bọc cách nhiệt HDPUF: KP39-W21; KP39-W27; KP39-W28; KP39-W30', 'unit': 'Joint', 'qty': 4, 'accum': 19}, {'stt': 9, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}], 'sec7': ['Kéo đẩy chuỗi ống xuống hào', 'Vệ sinh ống', 'Fit-up và hàn ống 30 inch', 'NDT mối hàn', 'Bắn cát dán FJC', 'Sắp xếp dọn dẹp bãi']},
    '01/11/2026': {'date': '01/11/2026', 'weather': 'Raining', 'worker': 34, 'sec4': [{'stt': 1, 'content': 'Tiếp nhận vật tư ống', 'unit': 'Pcs', 'qty': 16, 'accum': 80}, {'stt': 2, 'content': 'Kiểm tra ống đầu vào và vệ sinh làm sạch lòng ống', 'unit': 'Pcs', 'qty': 8, 'accum': 80}, {'stt': 3, 'content': 'Nghiệm thu Fit-up & hàn mối hàn: KP38-W80; KP39-W20; KP39-W22', 'unit': 'Joint', 'qty': 3, 'accum': 36}, {'stt': 4, 'content': 'Nghiệm thu Visual mối hàn: KP39-W07', 'unit': 'Joint', 'qty': 1, 'accum': 31}, {'stt': 5, 'content': 'Nghiệm thu NDT AUT mối hàn: KP39-W07; KP39-W09; KP39-W24; KP39-W26', 'unit': 'Joint', 'qty': 4, 'accum': 31}, {'stt': 6, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': 8}, {'stt': 7, 'content': 'Làm sạch bề mặt, dán FJC, thử nghiệm: KP39-W17; KP39-W13; KP39-W15; KP39-W24; KP39-W26', 'unit': 'Joint', 'qty': 5, 'accum': 28}, {'stt': 8, 'content': 'Nghiệm thu bọc cách nhiệt HDPUF: KP39-W28; KP39-W23; KP39-W24; KP39-W26; KP39-W25', 'unit': 'Joint', 'qty': 5, 'accum': 25}, {'stt': 9, 'content': 'Nghiệm thu kéo đẩy ống xuống hào: Ống 7253; 3401; 3782; 6958; 7698', 'unit': 'Ống', 'qty': 5, 'accum': 24}, {'stt': 10, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}], 'sec7': ['Kéo đẩy chuỗi ống xuống hào', 'Tiếp nhận và vệ sinh làm sạch lòng ống', 'Fit-up và hàn ống 30 inch', 'NDT mối hàn', 'Bắn cát, dán FJC ống', 'Sắp xếp dọn dẹp vệ sinh bãi']},
    '02/11/2026': {'date': '02/11/2026', 'weather': 'Sunny', 'worker': 34, 'sec4': [{'stt': 1, 'content': 'Nghiệm thu Fit-up & hàn mối hàn: KP39-W01; KP39-W03; KP39-W16; KP39-W18', 'unit': 'Joint', 'qty': 4, 'accum': 39}, {'stt': 2, 'content': 'Nghiệm thu Visual & NDT AUT: KP39-W05; KP39-W20; KP39-W22; KP38-W80', 'unit': 'Joint', 'qty': 4, 'accum': 35}, {'stt': 3, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': 9}, {'stt': 4, 'content': 'Làm sạch bề mặt, dán FJC, thử nghiệm: KP39-W09; KP39-W11; KP39-W20; KP39-W22', 'unit': 'Joint', 'qty': 4, 'accum': 32}, {'stt': 5, 'content': 'Nghiệm thu bọc cách nhiệt HDPUF: KP39-W19; KP39-W20; KP39-W21; KP39-W22', 'unit': 'Joint', 'qty': 4, 'accum': 27}, {'stt': 6, 'content': 'Nghiệm thu kéo đẩy ống xuống hào: KP39-W19; KP39-W20; KP39-W21; KP39-W22', 'unit': 'Joint', 'qty': 4, 'accum': 23}, {'stt': 7, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}], 'sec7': ['Kéo đẩy chuỗi ống xuống hào', 'Vệ sinh làm sạch lòng ống', 'Fit-up và hàn ống 30 inch', 'NDT mối hàn', 'Bắn cát, dán FJC ống', 'Sắp xếp dọn dẹp vệ sinh bãi']},
    '03/11/2026': {
        'date': '03/11/2026',
        'weather': 'Sunny',
        'worker': 33,
        'sec4': [
            {'stt': 1, 'content': 'Kiểm tra ống đầu vào và vệ sinh làm sạch lòng ống', 'unit': 'Pcs', 'qty': 8, 'accum': 88},
            {'stt': 2, 'content': 'Nghiệm thu Fit-up & hàn mối hàn: KP38-W76; KP38-W78; KP39-W12; KP39-W14', 'unit': 'Joint', 'qty': 4, 'accum': 43},
            {'stt': 3, 'content': 'Nghiệm thu Visual & NDT AUT mối hàn: KP39-W01; KP39-W03; KP39-W16; KP39-W18', 'unit': 'Joint', 'qty': 4, 'accum': 39},
            {'stt': 4, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': 10},
            {'stt': 5, 'content': 'Làm sạch bề mặt & dán FJC: KP38-W80; KP39-W01; KP39-W03; KP39-W05; KP39-W07; KP39-W16; KP39-W18', 'unit': 'Joint', 'qty': 7, 'accum': 39},
            {'stt': 6, 'content': 'Nghiệm thu bọc cách nhiệt HDPUF: KP39-W15; KP39-W16; KP39-W17; KP39-W18', 'unit': 'Joint', 'qty': 4, 'accum': 31},
            {'stt': 7, 'content': 'Nghiệm thu kéo đẩy ống xuống hào: KP39-W19; KP39-W20; KP39-W21; KP39-W22', 'unit': 'Joint', 'qty': 4, 'accum': 27},
            {'stt': 8, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}
        ],
        'sec7': [
            'Kéo đẩy chuỗi ống xuống hào',
            'Tiếp nhận vận chuyển ống vào bãi thi công',
            'Vệ sinh làm sạch lòng ống',
            'Fit-up và hàn ống 30 inch',
            'NDT mối hàn',
            'Bắn cát, dán FJC ống',
            'Sắp xếp dọn dẹp vệ sinh bãi'
        ]
    },
    '04/11/2026': {
        'date': '04/11/2026',
        'weather': 'Sunny',
        'worker': 34,
        'sec4': [
            {'stt': 1, 'content': 'Tiếp nhận vật tư ống', 'unit': 'Pcs', 'qty': 16, 'accum': 96},
            {'stt': 2, 'content': 'Nghiệm thu Fit-up & hàn mối hàn: KP38-W72; KP38-W74; KP39-W08; KP39-W10', 'unit': 'Joint', 'qty': 4, 'accum': 47},
            {'stt': 3, 'content': 'Nghiệm thu Visual & NDT AUT mối hàn: KP38-W76; KP38-W78; KP39-W12; KP39-W14', 'unit': 'Joint', 'qty': 4, 'accum': 43},
            {'stt': 4, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': 11},
            {'stt': 5, 'content': 'Làm sạch bề mặt & dán FJC: KP38-W76; KP38-W78; KP39-W12; KP39-W14', 'unit': 'Joint', 'qty': 4, 'accum': 43},
            {'stt': 6, 'content': 'Nghiệm thu bọc cách nhiệt HDPUF: KP39-W11; KP39-W12; KP39-W13; KP39-W14', 'unit': 'Joint', 'qty': 4, 'accum': 35},
            {'stt': 7, 'content': 'Nghiệm thu kéo đẩy ống xuống hào: KP39-W15; KP39-W16; KP39-W17; KP39-W18', 'unit': 'Joint', 'qty': 4, 'accum': 31},
            {'stt': 8, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}
        ],
        'sec7': [
            'Kéo đẩy chuỗi ống xuống hào',
            'Vệ sinh làm sạch lòng ống',
            'Fit-up và hàn ống 30 inch',
            'NDT mối hàn',
            'Bắn cát, dán FJC ống',
            'Sắp xếp dọn dẹp vệ sinh bãi'
        ]
    },
    '05/11/2026': {
        'date': '05/11/2026',
        'weather': 'Sunny',
        'worker': 35,
        'sec4': [
            {'stt': 1, 'content': 'Nghiệm thu Fit-up & hàn mối hàn: KP38-W64; KP38-W66; KP38-W68; KP39-W04; KP39-W06', 'unit': 'Joint', 'qty': 5, 'accum': 52},
            {'stt': 2, 'content': 'Nghiệm thu Visual & NDT AUT mối hàn: KP38-W72; KP38-W74; KP39-W08; KP39-W10', 'unit': 'Joint', 'qty': 4, 'accum': 47},
            {'stt': 3, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': 12},
            {'stt': 4, 'content': 'Làm sạch bề mặt & dán FJC: KP38-W72; KP38-W74; KP39-W08; KP39-W10', 'unit': 'Joint', 'qty': 4, 'accum': 47},
            {'stt': 5, 'content': 'Nghiệm thu bọc cách nhiệt HDPUF: KP39-W07; KP39-W08; KP39-W09; KP39-W10', 'unit': 'Joint', 'qty': 4, 'accum': 39},
            {'stt': 6, 'content': 'Nghiệm thu kéo đẩy ống xuống hào: KP39-W11; KP39-W12; KP39-W13; KP39-W14', 'unit': 'Joint', 'qty': 4, 'accum': 35},
            {'stt': 7, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}
        ],
        'sec7': [
            'Kéo đẩy chuỗi ống xuống hào',
            'Vệ sinh làm sạch lòng ống',
            'Fit-up và hàn ống 30 inch',
            'NDT mối hàn',
            'Bắn cát, dán FJC ống',
            'Sắp xếp dọn dẹp vệ sinh bãi'
        ]
    },
    '06/11/2026': {
        'date': '06/11/2026',
        'weather': 'Sunny',
        'worker': 34,
        'sec4': [
            {'stt': 1, 'content': 'Nghiệm thu tiếp nhận vật liệu hàn', 'unit': 'Lot', 'qty': 1, 'accum': 2},
            {'stt': 2, 'content': 'Nghiệm thu Fit-up & hàn mối hàn: KP38-W56; KP38-W58; KP38-W70; KP38-W81; KP39-W02', 'unit': 'Joint', 'qty': 5, 'accum': 57},
            {'stt': 3, 'content': 'Nghiệm thu Visual & NDT AUT mối hàn: KP38-W64; KP38-W66; KP39-W04; KP39-W06', 'unit': 'Joint', 'qty': 4, 'accum': 51},
            {'stt': 4, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': 13},
            {'stt': 5, 'content': 'Làm sạch bề mặt & dán FJC: KP38-W64; KP38-W66; KP39-W04; KP39-W06', 'unit': 'Joint', 'qty': 4, 'accum': 51},
            {'stt': 6, 'content': 'Nghiệm thu bọc cách nhiệt HDPUF: KP39-W03; KP39-W04; KP39-W05; KP39-W06', 'unit': 'Joint', 'qty': 4, 'accum': 43},
            {'stt': 7, 'content': 'Nghiệm thu kéo đẩy ống xuống hào: KP39-W07; KP39-W08; KP39-W09; KP39-W10', 'unit': 'Joint', 'qty': 4, 'accum': 39},
            {'stt': 8, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}
        ],
        'sec7': [
            'Tiếp nhận vận chuyển ống vào bãi thi công',
            'Kéo đẩy chuỗi ống xuống hào',
            'Vệ sinh làm sạch lòng ống',
            'Fit-up và hàn ống 30 inch',
            'NDT mối hàn',
            'Bắn cát, dán FJC ống',
            'Sắp xếp dọn dẹp vệ sinh bãi'
        ]
    },
    '07/11/2026': {
        'date': '07/11/2026',
        'weather': 'Sunny',
        'worker': 35,
        'sec4': [
            {'stt': 1, 'content': 'Nghiệm thu Fit-up & hàn mối hàn: KP38-W54; KP38-W60; KP38-W62; KP38-W77; KP38-W79', 'unit': 'Joint', 'qty': 5, 'accum': 62},
            {'stt': 2, 'content': 'Nghiệm thu Visual & NDT AUT mối hàn: KP38-W56; KP38-W58; KP38-W68; KP38-W70; KP38-W81; KP39-W02', 'unit': 'Joint', 'qty': 6, 'accum': 57},
            {'stt': 3, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': 14},
            {'stt': 4, 'content': 'Làm sạch bề mặt & dán FJC: KP38-W56; KP38-W58; KP38-W81; KP39-W02', 'unit': 'Joint', 'qty': 4, 'accum': 55},
            {'stt': 5, 'content': 'Nghiệm thu bọc cách nhiệt HDPUF: KP38-W80; KP38-W81; KP39-W01; KP39-W02', 'unit': 'Joint', 'qty': 4, 'accum': 47},
            {'stt': 6, 'content': 'Nghiệm thu kéo đẩy ống xuống hào: KP39-W03; KP39-W04; KP39-W05; KP39-W06', 'unit': 'Joint', 'qty': 4, 'accum': 43},
            {'stt': 7, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''}
        ],
        'sec7': [
            'Kéo đẩy chuỗi ống xuống hào',
            'Vệ sinh làm sạch lòng ống',
            'Fit-up và hàn ống 30 inch',
            'NDT mối hàn',
            'Bắn cát, dán FJC ống',
            'Sắp xếp dọn dẹp vệ sinh bãi'
        ]
    },
}

def auto_discover_new_days(ctrl_xlsx, form_xlsx, known_days):
    """Tự động quét và tổng hợp chính xác các ngày mới từ Controlog WM (sheet 'Hoàn thành Joint') và FORM CHOOSE REPORT"""
    new_days = dict(known_days)
    try:
        import openpyxl
        
        # 1. Thử đọc từ sheet 'Hoàn thành Joint' trong file Controlog nếu có
        ht_data_by_date = {}
        all_unique_dates = set()
        if ctrl_xlsx and os.path.exists(ctrl_xlsx):
            try:
                wb_ctrl = openpyxl.load_workbook(ctrl_xlsx, data_only=True)
                if 'Hoàn thành Joint' in wb_ctrl.sheetnames:
                    ws_ht = wb_ctrl['Hoàn thành Joint']
                    steps_cols = [
                        ('fitup', 5, 6),
                        ('visual', 9, 10),
                        ('ndt', 11, 12),
                        ('fjc', 13, 14),
                        ('hti', 19, 20),
                        ('lowering', 23, 24),
                    ]
                    
                    # Thu thập từng bước theo từng ngày
                    date_joints = {} # date_obj -> step -> list of joints
                    for r in range(7, ws_ht.max_row + 1):
                        j_name = ws_ht.cell(r, 3).value
                        if not j_name:
                            continue
                        for s_key, c_dt, c_rp in steps_cols:
                            dt_val = ws_ht.cell(r, c_dt).value
                            if isinstance(dt_val, (datetime.date, datetime.datetime)):
                                d_obj = dt_val.date() if isinstance(dt_val, datetime.datetime) else dt_val
                                all_unique_dates.add(d_obj)
                                date_joints.setdefault(d_obj, {}).setdefault(s_key, []).append(str(j_name))
                                
                    # Tính toán lũy kế chạy cho từng ngày
                    sorted_d_objs = sorted(list(all_unique_dates))
                    running_acc = {s_key: 0 for s_key, _, _ in steps_cols}
                    for d_obj in sorted_d_objs:
                        d_str = d_obj.strftime('%d/%m/%Y')
                        day_counts = {}
                        for s_key, _, _ in steps_cols:
                            j_list = date_joints.get(d_obj, {}).get(s_key, [])
                            cnt = len(j_list)
                            running_acc[s_key] += cnt
                            day_counts[s_key] = {
                                'joints': j_list,
                                'qty': cnt,
                                'accum': running_acc[s_key]
                            }
                        ht_data_by_date[d_str] = day_counts
                wb_ctrl.close()
            except Exception as e_ht:
                print(f"  [Lưu ý] Lỗi đọc sheet Hoàn thành Joint: {e_ht}")

        # 2. Đọc bổ sung các thông tin ngoài Joint từ file FORM CHOOSE REPORT (ICR, MRIR, DIC)
        extra_by_date = {}
        if form_xlsx and os.path.exists(form_xlsx):
            try:
                wb_form = openpyxl.load_workbook(form_xlsx, data_only=True)
                if 'FORM' in wb_form.sheetnames:
                    s_form = wb_form['FORM']
                    for r in range(4, s_form.max_row + 1):
                        dt = s_form.cell(r, 14).value
                        if isinstance(dt, (datetime.date, datetime.datetime)) and dt.year >= 2026:
                            d_str = dt.strftime('%d/%m/%Y')
                            code = str(s_form.cell(r, 8).value or '').upper().strip()
                            desc = s_form.cell(r, 12).value or ''
                            extra_by_date.setdefault(d_str, []).append((code, desc))
                wb_form.close()
            except Exception as e_form:
                print(f"  [Lưu ý] Lỗi đọc sheet FORM: {e_form}")

        # 3. Lấy tất cả ngày xuất hiện cần bổ sung
        all_candidate_dates = set(ht_data_by_date.keys()).union(extra_by_date.keys())
        
        # Khởi tạo running accumulators từ các ngày đã biết (new_days)
        running_pipe_mrir = 0
        running_pipe_icr = 0
        running_weld_mrir = 0
        running_dic = 0
        for d_k, d_val in new_days.items():
            for it in d_val.get('sec4', []):
                c_low = it.get('content', '').lower()
                acc = it.get('accum')
                if isinstance(acc, int) or (isinstance(acc, str) and acc.isdigit()):
                    val_acc = int(acc)
                    if 'tiếp nhận vật tư ống' in c_low:
                        running_pipe_mrir = max(running_pipe_mrir, val_acc)
                    elif 'vệ sinh làm sạch lòng ống' in c_low or 'kiểm tra ống đầu vào' in c_low:
                        running_pipe_icr = max(running_pipe_icr, val_acc)
                    elif 'tiếp nhận vật liệu hàn' in c_low:
                        running_weld_mrir = max(running_weld_mrir, val_acc)
                    elif 'hiệu chuẩn' in c_low:
                        running_dic = max(running_dic, val_acc)

        for d_str in sorted(all_candidate_dates, key=lambda x: datetime.datetime.strptime(x, '%d/%m/%Y')):
            if d_str in new_days:
                continue
                
            print(f"  [Tự động nhận diện ngày mới]: {d_str}")
            sec4 = []
            stt = 1
            
            # A. Kiểm tra ICR / MRIR
            ext_items = extra_by_date.get(d_str, [])
            for code, desc in ext_items:
                if code == 'MRIR' and 'WELDING' in desc.upper():
                    running_weld_mrir += 1
                    sec4.append({'stt': stt, 'content': 'Nghiệm thu tiếp nhận vật liệu hàn', 'unit': 'Lot', 'qty': 1, 'accum': running_weld_mrir})
                    stt += 1
                elif code == 'MRIR' and ('PIPE' in desc.upper() or 'VẬT TƯ' in desc.upper() or 'ỐNG' in desc.upper()):
                    qm = re.search(r'(\d+)\s*(?:pcs|ống)', desc, re.I)
                    q_val = int(qm.group(1)) if qm else 16
                    running_pipe_mrir += q_val
                    sec4.append({'stt': stt, 'content': 'Tiếp nhận vật tư ống', 'unit': 'Pcs', 'qty': q_val, 'accum': running_pipe_mrir})
                    stt += 1
                elif code == 'ICR':
                    qm = re.search(r'(\d+)\s*(?:pcs|ống)', desc, re.I)
                    q_val = int(qm.group(1)) if qm else 8
                    running_pipe_icr += q_val
                    sec4.append({'stt': stt, 'content': 'Kiểm tra ống đầu vào và vệ sinh làm sạch lòng ống', 'unit': 'Pcs', 'qty': q_val, 'accum': running_pipe_icr})
                    stt += 1

            # B. Các bước thi công mối hàn từ bảng Hoàn thành Joint
            ht_day = ht_data_by_date.get(d_str, {})
            if ht_day.get('fitup') and ht_day['fitup']['qty'] > 0:
                js = '; '.join(ht_day['fitup']['joints'])
                sec4.append({'stt': stt, 'content': f'Nghiệm thu Fit-up & hàn mối hàn: {js}', 'unit': 'Joint', 'qty': ht_day['fitup']['qty'], 'accum': ht_day['fitup']['accum']})
                stt += 1
            if ht_day.get('ndt') and ht_day['ndt']['qty'] > 0:
                js = '; '.join(ht_day['ndt']['joints'])
                sec4.append({'stt': stt, 'content': f'Nghiệm thu Visual & NDT AUT mối hàn: {js}', 'unit': 'Joint', 'qty': ht_day['ndt']['qty'], 'accum': ht_day['ndt']['accum']})
                stt += 1
                
            # Hiệu chuẩn thiết bị đo
            running_dic += 1
            sec4.append({'stt': stt, 'content': 'Hiệu chuẩn dụng cụ đo', 'unit': 'Bộ', 'qty': 1, 'accum': running_dic})
            stt += 1
            
            if ht_day.get('fjc') and ht_day['fjc']['qty'] > 0:
                js = '; '.join(ht_day['fjc']['joints'])
                sec4.append({'stt': stt, 'content': f'Làm sạch bề mặt & dán FJC: {js}', 'unit': 'Joint', 'qty': ht_day['fjc']['qty'], 'accum': ht_day['fjc']['accum']})
                stt += 1
            if ht_day.get('hti') and ht_day['hti']['qty'] > 0:
                js = '; '.join(ht_day['hti']['joints'])
                sec4.append({'stt': stt, 'content': f'Nghiệm thu bọc cách nhiệt HDPUF: {js}', 'unit': 'Joint', 'qty': ht_day['hti']['qty'], 'accum': ht_day['hti']['accum']})
                stt += 1
            if ht_day.get('lowering') and ht_day['lowering']['qty'] > 0:
                js = '; '.join(ht_day['lowering']['joints'])
                sec4.append({'stt': stt, 'content': f'Nghiệm thu kéo đẩy ống xuống hào: {js}', 'unit': 'Joint', 'qty': ht_day['lowering']['qty'], 'accum': ht_day['lowering']['accum']})
                stt += 1
                
            sec4.append({'stt': stt, 'content': 'Sắp xếp dọn dẹp vệ sinh bãi thi công', 'unit': '', 'qty': '', 'accum': ''})
            
            new_days[d_str] = {
                'date': d_str,
                'weather': 'Sunny',
                'worker': 34,
                'sec4': sec4,
                'sec7': [
                    'Kéo đẩy chuỗi ống xuống hào',
                    'Vệ sinh làm sạch lòng ống',
                    'Fit-up và hàn ống 30 inch',
                    'NDT mối hàn',
                    'Bắn cát, dán FJC ống',
                    'Sắp xếp dọn dẹp vệ sinh bãi'
                ]
            }
    except Exception as e:
        print(f"  [Lưu ý] Không đọc được dữ liệu Excel bổ sung: {e}")
        
    return new_days

def apply_clean_xml_fixes(xml_str, orig_root_tag):
    # 1. Bảo tồn 100% thẻ root <w:document ...> gốc
    curr_start = xml_str.find('<w:document')
    curr_end = xml_str.find('>', curr_start)
    xml_str = xml_str[:curr_start] + orig_root_tag[orig_root_tag.find('<w:document'):] + xml_str[curr_end+1:]
    
    # 2. Khôi phục khai báo cục bộ bắt buộc cho DrawingML
    xml_str = xml_str.replace('<a:graphic>', '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">')
    xml_str = xml_str.replace('<pic:pic>', '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">')
    xml_str = xml_str.replace('<a:graphicFrameLocks noChangeAspect="1" />', '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>')
    xml_str = xml_str.replace('<a:graphicFrameLocks noChangeAspect="1"/>', '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>')
    xml_str = xml_str.replace('<a14:useLocalDpi val="0" />', '<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>')
    xml_str = xml_str.replace('<a14:useLocalDpi val="0"/>', '<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>')
    
    # 3. Đảm bảo tất cả các section (sectPr) đều có footerReference để hiển thị từ trang đầu tới hết
    def add_fr(match):
        content = match.group(0)
        if 'w:footerReference' not in content:
            return content.replace('</w:sectPr>', '<w:footerReference w:type="default" r:id="rId28"/></w:sectPr>')
        return content
        
    xml_str = re.sub(r'<w:sectPr[^>]*>.*?</w:sectPr>', add_fr, xml_str, flags=re.DOTALL)
    return xml_str

def update_template_footer(template_docx):
    """Cập nhật footer vào ngay trong file mẫu 1. Nhật ký thi công (R1).docx"""
    if not os.path.exists(template_docx):
        return
    try:
        with zipfile.ZipFile(template_docx, 'r') as zin:
            file_map = {item.filename: zin.read(item.filename) for item in zin.infolist()}
            
        footer_bytes = FOOTER_XML_TEMPLATE.encode('utf-8')
        for f_name in ['word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml', 'word/footer4.xml']:
            file_map[f_name] = footer_bytes
            
        orig_doc_xml = file_map['word/document.xml'].decode('utf-8')
        def add_fr(match):
            content = match.group(0)
            if 'w:footerReference' not in content:
                return content.replace('</w:sectPr>', '<w:footerReference w:type="default" r:id="rId28"/></w:sectPr>')
            return content
            
        new_doc_xml = re.sub(r'<w:sectPr[^>]*>.*?</w:sectPr>', add_fr, orig_doc_xml, flags=re.DOTALL)
        file_map['word/document.xml'] = new_doc_xml.encode('utf-8')
        
        with zipfile.ZipFile(template_docx, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for fname, data in file_map.items():
                zout.writestr(fname, data)
        print(f"  [OK] Đã tích hợp Footer tự động vào file mẫu: {os.path.basename(template_docx)}")
    except Exception as e:
        print(f"  [Lưu ý] Không thể cập nhật footer vào file mẫu: {e}")

def generate_monthly_docs(template_docx, output_folder, days_data_by_month):
    """Tạo các file Word nhật ký thi công theo từng tháng"""
    if not os.path.exists(template_docx):
        print(f"Lỗi: Không tìm thấy file mẫu Word: {template_docx}")
        return []
        
    with zipfile.ZipFile(template_docx, 'r') as zin:
        file_map = {item.filename: zin.read(item.filename) for item in zin.infolist()}

    # Nạp và cập nhật toàn bộ các file footer (footer1..footer4)
    footer_bytes = FOOTER_XML_TEMPLATE.encode('utf-8')
    for f_name in ['word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml', 'word/footer4.xml']:
        file_map[f_name] = footer_bytes

    orig_doc_xml = file_map['word/document.xml'].decode('utf-8')
    orig_root_tag = orig_doc_xml[:orig_doc_xml.find('>', orig_doc_xml.find('<w:document')) + 1]
    
    root = ET.fromstring(orig_doc_xml)
    body = root.find('w:body', ns)
    
    # Front matter: 13 phần tử đầu tiên (gồm 4 bảng bìa & info)
    front_matter = [copy.deepcopy(body[i]) for i in range(13)]
    
    # Daily templates
    tblA_template = copy.deepcopy(body[43])
    p_pagebreak = copy.deepcopy(body[44]) # br page
    tblB_template = copy.deepcopy(body[45])
    p_mid = copy.deepcopy(body[46]) # mid spacing p
    tblC_template = copy.deepcopy(body[47])
    p_daybreak = copy.deepcopy(body[42]) # day br page
    p_final = copy.deepcopy(body[55]) # final p
    sectPr = copy.deepcopy(body[56]) # final sectPr
    if sectPr.find('w:footerReference', ns) is None:
        fr = ET.SubElement(sectPr, f'{{{w_ns}}}footerReference')
        fr.attrib[f'{{{w_ns}}}type'] = 'default'
        fr.attrib[f'{{{r_ns}}}id'] = 'rId28'

    os.makedirs(output_folder, exist_ok=True)
    generated_files = []

    for ym_key in sorted(days_data_by_month.keys()):
        days_data = days_data_by_month[ym_key]
        year, month = ym_key.split('-')
        month_int = int(month)
        
        # Tên file theo định dạng chuẩn: 2. NKTC Tháng 10 - TTC11.docx, 3. NKTC Tháng 11 - TTC11.docx
        stt_prefix = 2 + (month_int - 10) if month_int >= 10 else month_int
        file_name = f"{stt_prefix}. NKTC Tháng {month_int} - TTC11.docx"
        output_docx = os.path.join(output_folder, file_name)
        
        print(f"\n=======================================================")
        print(f" Đang tạo: {file_name} ({len(days_data)} ngày thi công)")
        print(f"=======================================================")
        
        new_body = copy.deepcopy(body)
        new_body.clear()
        
        for elem in front_matter:
            new_body.append(elem)
            
        for day_idx, d_data in enumerate(days_data):
            w_disp = d_data.get('weather')
            cn_disp = d_data.get('worker') if d_data.get('worker') is not None else "(chưa có)"
            print(f"  -> Thêm ngày {day_idx+1}/{len(days_data)}: {d_data['date']} (Thời tiết: {w_disp}, CN: {cn_disp})")
            tblA = build_tblA(tblA_template, d_data)
            new_body.append(tblA)
            new_body.append(copy.deepcopy(p_pagebreak))
            
            tblB = build_tblB(tblB_template, d_data)
            new_body.append(tblB)
            new_body.append(copy.deepcopy(p_mid))
            
            tblC = build_tblC(tblC_template, d_data)
            new_body.append(tblC)
            
            if day_idx < len(days_data) - 1:
                new_body.append(copy.deepcopy(p_daybreak))
            else:
                new_body.append(copy.deepcopy(p_final))
                new_body.append(copy.deepcopy(sectPr))
                
        root_copy = copy.deepcopy(root)
        old_body = root_copy.find('w:body', ns)
        root_copy.remove(old_body)
        root_copy.append(new_body)
        
        raw_new_xml = ET.tostring(root_copy, encoding='utf-8', xml_declaration=True).decode('utf-8')
        clean_doc_xml = apply_clean_xml_fixes(raw_new_xml, orig_root_tag)
        file_map_copy = dict(file_map)
        file_map_copy['word/document.xml'] = clean_doc_xml.encode('utf-8')
        
        try:
            with zipfile.ZipFile(output_docx, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for fname, data in file_map_copy.items():
                    zout.writestr(fname, data)
            sz = os.path.getsize(output_docx)
            print(f"  [Thành công] Đã tạo file: {output_docx} ({sz:,} bytes)")
            generated_files.append(output_docx)
        except PermissionError:
            print(f"\n[CẢNH BÁO LỖI] Không thể ghi đè file: '{output_docx}'!")
            print("Nguyên nhân: File này đang được mở trong Microsoft Word.")
            print("Vui lòng ĐÓNG FILE WORD ĐANG MỞ và chạy lại công cụ.")
            
    return generated_files

def get_workspace_dir():
    """Tự động xác định thư mục chứa dữ liệu (file mẫu Word và Excel)"""
    if getattr(sys, 'frozen', False):
        app_dir = os.path.dirname(sys.executable)
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))
        
    # Ưu tiên thư mục làm việc hiện tại
    cwd = os.getcwd()
    for candidate in [cwd, app_dir, os.path.dirname(app_dir), os.path.dirname(os.path.dirname(app_dir))]:
        if glob.glob(os.path.join(candidate, "1. Nhật ký thi công*.docx")):
            return candidate
            
    return app_dir

def run_pipeline():
    workspace_dir = get_workspace_dir()
    os.chdir(workspace_dir)
    
    # 1. File đầu vào
    template_docx = find_file("1. Nhật ký thi công*.docx", "1. Nhật ký thi công (R1).docx")
    ctrl_xlsx = find_file("Controlog*.xlsx", "Controlog WM TTC11.xlsx")
    form_xlsx = find_file("FORM CHOOSE REPORT*.xlsx", "FORM CHOOSE REPORT Rev 8_ORIGINAL- TTC11.xlsx")
    
    # Cập nhật footer vào file mẫu nếu cần
    update_template_footer(template_docx)

    print(f"\n[1] Kiểm tra file dữ liệu:")
    print(f"  - File mẫu Word : {os.path.basename(template_docx)} ({'OK' if os.path.exists(template_docx) else 'THIẾU'})")
    print(f"  - Control Log   : {os.path.basename(ctrl_xlsx)} ({'OK' if os.path.exists(ctrl_xlsx) else 'THIẾU'})")
    print(f"  - FORM Report   : {os.path.basename(form_xlsx)} ({'OK' if os.path.exists(form_xlsx) else 'THIẾU'})")
    
    if not os.path.exists(template_docx):
        print(f"\n[Lỗi] Thiếu file mẫu Word '{template_docx}'. Vui lòng đặt file mẫu vào thư mục tool.")
        return []

    # 2. Nạp và mở rộng ngày nếu có ngày mới trong Excel
    all_days = auto_discover_new_days(ctrl_xlsx, form_xlsx, KNOWN_DAYS_DATA)

    # 3. Quét Zalo theo dải ngày thực tế
    zalo_candidates = [
        os.path.join(workspace_dir, "HinhAnh_BaoCao"),
        os.path.join(workspace_dir, "anh_zalo"),
        os.path.join(workspace_dir, "Zalo"),
        r"D:\ZaloData\media\1506833805674878003\ZaloDownloads\resource\g5241664229805465444\picture"
    ]
    zalo_dir = None
    for zc in zalo_candidates:
        if os.path.isdir(zc):
            zalo_dir = zc
            break
            
    zalo_reports = {}
    if zalo_dir:
        print(f"  - Thư mục ảnh Zalo: {zalo_dir}")
        zalo_reports = scan_zalo_reports(zalo_dir, target_dates=list(all_days.keys()))
    else:
        print("  - Thư mục ảnh Zalo: Không tìm thấy (Sẽ dùng thời tiết 'Bình thường' mặc định).")
        print("    (Mẹo: Bạn có thể tạo thư mục 'HinhAnh_BaoCao' và chép ảnh Zalo vào).")
    
    # Áp dụng thời tiết và nhân lực từ Zalo nếu có
    for d_str, d_val in all_days.items():
        if d_str in zalo_reports:
            if zalo_reports[d_str].get('weather'):
                d_val['weather'] = zalo_reports[d_str]['weather']
            if zalo_reports[d_str].get('worker') is not None:
                d_val['worker'] = zalo_reports[d_str]['worker']

    # Nhóm theo tháng YYYY-MM
    days_by_month = {}
    for d_str in sorted(all_days.keys(), key=lambda x: datetime.datetime.strptime(x, '%d/%m/%Y')):
        dt = datetime.datetime.strptime(d_str, '%d/%m/%Y')
        ym_key = dt.strftime('%Y-%m')
        days_by_month.setdefault(ym_key, []).append(all_days[d_str])

    # 4. Sinh file Word
    results = generate_monthly_docs(template_docx, workspace_dir, days_by_month)
    return results

def main():
    print("=" * 65)
    print("      CÔNG CỤ TỰ ĐỘNG TẠO NHẬT KÝ THI CÔNG - BÃI TTC11")
    print("         Dự án: Đường ống dẫn khí Lô B - Ô Môn")
    print("=" * 65)
    
    results = run_pipeline()
    
    print("\n" + "=" * 65)
    print("               HOÀN THÀNH TẤT CẢ CÔNG VIỆC")
    print("=" * 65)
    for rf in results:
        print(f"  [OK] {os.path.basename(rf)}")
    print("\nFile nhật ký thi công đã được tạo thành công.")
    try:
        input("\nNhấn Enter để kết thúc...")
    except (EOFError, KeyboardInterrupt):
        pass

if __name__ == '__main__':
    main()


# ==============================================================================
# INTEGRATION API FOR WEB APPLICATION (FOME_WEB)
# ==============================================================================
import io
from pathlib import Path

def get_site_nktc_template_path(location: str) -> Path:
    site_p = Path(f"Fome_Sites/{location}/Fome_Web_Data")
    if site_p.is_dir():
        for cand in site_p.glob("*.docx"):
            if "nhat ky" in cand.name.lower() or "nktc" in cand.name.lower():
                return cand
    for cand in [Path("Fome_Web_Data/Templates/1. Nhật ký thi công (R1).docx"),
                 Path("1. Nhật ký thi công (R1).docx")]:
        if cand.is_file():
            return cand
    return Path(f"Fome_Sites/{location}/Fome_Web_Data/1. Nhật ký thi công {location}.docx")

def get_site_nktc_template_info(location: str) -> dict:
    p = get_site_nktc_template_path(location)
    if p and p.is_file():
        stat = p.stat()
        return {
            "exists": True,
            "filename": p.name,
            "size": stat.st_size,
            "modified": datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%d/%m/%Y %H:%M:%S")
        }
    return {
        "exists": False,
        "filename": "Chưa nạp file mẫu Word",
        "size": 0,
        "modified": ""
    }

def save_site_nktc_template(location: str, filename: str, content: bytes) -> dict:
    if not filename.lower().endswith(".docx"):
        raise ValueError("File mẫu Nhật ký thi công phải có định dạng Word (.docx).")
    try:
        with zipfile.ZipFile(io.BytesIO(content), "r") as z:
            if "word/document.xml" not in z.namelist():
                raise ValueError("Tệp không phải là file Word (.docx) hợp lệ.")
    except Exception as exc:
        raise ValueError(f"Không thể đọc file Word: {exc}")
        
    dest_dir = Path(f"Fome_Sites/{location}/Fome_Web_Data")
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest_path = dest_dir / f"1. Nhật ký thi công {location}.docx"
    dest_path.write_bytes(content)
    return get_site_nktc_template_info(location)

import json

_ZALO_REPORTS_CACHE = {}

def load_persistent_zalo_cache() -> dict:
    global _ZALO_REPORTS_CACHE
    candidates = [
        Path("Fome_Sites/TTC11/Fome_Web_Data/zalo_reports_cache.json"),
        Path("Fome_Web_Data/zalo_reports_cache.json"),
        Path("zalo_reports_cache.json")
    ]
    for cp in candidates:
        if cp.is_file():
            try:
                with open(cp, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if "GLOBAL" not in _ZALO_REPORTS_CACHE:
                        _ZALO_REPORTS_CACHE["GLOBAL"] = {}
                    _ZALO_REPORTS_CACHE["GLOBAL"].update(data)
            except Exception:
                pass
    return _ZALO_REPORTS_CACHE.get("GLOBAL", {})

def save_persistent_zalo_cache(data: dict, location: str = "TTC11"):
    cp = Path(f"Fome_Sites/{location}/Fome_Web_Data/zalo_reports_cache.json")
    try:
        cp.parent.mkdir(parents=True, exist_ok=True)
        cur = {}
        if cp.is_file():
            with open(cp, "r", encoding="utf-8") as f:
                cur = json.load(f)
        cur.update(data)
        with open(cp, "w", encoding="utf-8") as f:
            json.dump(cur, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def get_cached_zalo_reports(zalo_dir: str, target_dates: list = None, force_refresh: bool = False, location: str = "TTC11") -> dict:
    global _ZALO_REPORTS_CACHE
    cached_global = load_persistent_zalo_cache()
    if not force_refresh:
        # Nếu đã có đầy đủ ngày cần tìm trong cache, trả về ngay lập tức không cần OCR
        if target_dates:
            missing = [d for d in target_dates if d not in cached_global]
            if not missing:
                return cached_global
            # Chỉ quét đúng các ngày còn thiếu trong Zalo!
            target_dates = missing
        elif len(cached_global) >= 10 and (not zalo_dir or not os.path.isdir(zalo_dir)):
            return cached_global
            
    reports = scan_zalo_reports(zalo_dir, target_dates=target_dates)
    cached_global.update(reports)
    # Đánh dấu những ngày đã quét nhưng không có ảnh Zalo để không quét lại lặp đi lặp lại
    if target_dates:
        for d in target_dates:
            if d not in cached_global:
                cached_global[d] = {"weather": "Normal", "worker": None}
    _ZALO_REPORTS_CACHE["GLOBAL"] = cached_global
    save_persistent_zalo_cache(cached_global, location)
    return cached_global

def get_all_site_days(location: str) -> tuple[dict, list, list]:
    site_data = Path(f"Fome_Sites/{location}/Fome_Web_Data")
    ctrl_path = None
    form_path = None
    if site_data.is_dir():
        for cand in site_data.glob("*.xlsx"):
            if "controlog" in cand.name.lower():
                ctrl_path = cand
            elif "form" in cand.name.lower():
                form_path = cand
    if not ctrl_path:
        for cand in Path("Fome_Web_Data").glob(f"*{location}*.xlsx"):
            if "controlog" in cand.name.lower():
                ctrl_path = cand
    if not form_path:
        for cand in Path("Fome_Web_Data").glob(f"*{location}*.xlsx"):
            if "form" in cand.name.lower():
                form_path = cand

    base_known = KNOWN_DAYS_DATA if location.upper() == "TTC11" else {}
    all_days = auto_discover_new_days(str(ctrl_path) if ctrl_path else "", str(form_path) if form_path else "", base_known)
    
    # Try Zalo scanning if folder exists
    zalo_candidates = [
        Path(f"Fome_Sites/{location}/Fome_Web_Data/HinhAnh_BaoCao"),
        Path(f"Fome_Sites/{location}/HinhAnh_BaoCao"),
        Path("HinhAnh_BaoCao"),
        Path("D:/ZaloData/media/1506833805674878003/ZaloDownloads/resource/g5241664229805465444/picture")
    ]
    for zc in zalo_candidates:
        if zc.is_dir():
            try:
                target_dates_list = list(all_days.keys())
                z_reports = get_cached_zalo_reports(str(zc), target_dates=target_dates_list)
                for d_str, d_val in all_days.items():
                    if d_str in z_reports:
                        if z_reports[d_str].get("weather"):
                            d_val["weather"] = z_reports[d_str]["weather"]
                        if z_reports[d_str].get("worker") is not None:
                            d_val["worker"] = z_reports[d_str]["worker"]
            except Exception:
                pass
            break

    days_list = []
    months_set = set()
    for d_str in sorted(all_days.keys(), key=lambda x: datetime.datetime.strptime(x, "%d/%m/%Y")):
        dt = datetime.datetime.strptime(d_str, "%d/%m/%Y")
        ym = dt.strftime("%Y-%m")
        m_label = dt.strftime("%m/%Y")
        months_set.add((ym, m_label))
        d_val = all_days[d_str]
        days_list.append({
            "date": d_str,
            "year_month": ym,
            "month_label": m_label,
            "weather": d_val.get("weather", "Normal"),
            "worker": d_val.get("worker"),
            "items_count": len(d_val.get("sec4", []))
        })
    
    sorted_months = sorted(list(months_set), key=lambda x: x[0])
    return all_days, days_list, sorted_months

def export_nktc_document(location: str, selected_dates: list[str]) -> tuple[Path, str]:
    all_days, _, _ = get_all_site_days(location)
    template_path = get_site_nktc_template_path(location)
    if not template_path.is_file():
        raise ValueError(f"Không tìm thấy file mẫu Word Nhật ký thi công cho trạm {location}.")

    selected_set = set(selected_dates)
    filtered_days = {d: all_days[d] for d in selected_dates if d in all_days}
    if not filtered_days:
        filtered_days = all_days

    days_by_month = {}
    for d_str in sorted(filtered_days.keys(), key=lambda x: datetime.datetime.strptime(x, "%d/%m/%Y")):
        dt = datetime.datetime.strptime(d_str, "%d/%m/%Y")
        ym = dt.strftime("%Y-%m")
        days_by_month.setdefault(ym, []).append(filtered_days[d_str])

    output_dir = Path("Fome_Web_Output")
    output_dir.mkdir(parents=True, exist_ok=True)
    generated = generate_monthly_docs(str(template_path), str(output_dir), days_by_month)

    if not generated:
        raise ValueError("Không thể tạo file Word Nhật ký thi công.")

    if len(generated) == 1:
        p = Path(generated[0])
        return p, p.name

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_name = f"NKTC_{location}_TongHop_{timestamp}.zip"
    zip_path = output_dir / zip_name
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in generated:
            zf.write(f, Path(f).name)
    return zip_path, zip_name
