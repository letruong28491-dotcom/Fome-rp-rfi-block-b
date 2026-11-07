from __future__ import annotations

import copy

import html
import gc
import io
import time

import hashlib

import json

import math

import os

import posixpath

import random

import re

import shutil
import tempfile

import sys

import threading

import webbrowser

import zipfile

from concurrent.futures import ThreadPoolExecutor

from datetime import date, datetime

from dataclasses import dataclass

from email import policy

from email.parser import BytesParser

from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

try:
    import nktc_processor
    HAS_NKTC = True
except Exception as _nktc_err:
    HAS_NKTC = False
    print(f"[NKTC] Warning: Cannot import nktc_processor: {_nktc_err}")


from pathlib import Path

from urllib.parse import parse_qs, quote, unquote, urlparse

from xml.etree import ElementTree

from openpyxl import load_workbook

from openpyxl.cell.cell import MergedCell

from openpyxl.cell.rich_text import TextBlock, CellRichText

from openpyxl.cell.text import InlineFont

import openpyxl.cell.rich_text as _rt

def _safe_whitespace(node):

    if node.text and (node.text.startswith(' ') or node.text.endswith(' ') or node.text.strip() == ''):

        node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

_rt.whitespace = _safe_whitespace

from openpyxl.formula.translate import Translator

from openpyxl.styles import Border, Side, Font, Alignment

from openpyxl.utils.cell import column_index_from_string, get_column_letter

from openpyxl.worksheet.pagebreak import Break

if sys.stdout and hasattr(sys.stdout, "reconfigure"):

    try:

        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    except Exception:

        pass

if sys.stderr and hasattr(sys.stderr, "reconfigure"):

    try:

        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    except Exception:

        pass

if getattr(sys, "frozen", False):

    ROOT = Path(sys.executable).resolve().parent

    if not (ROOT / "RFI 15A").exists() and (ROOT.parent.parent / "RFI 15A").exists():

        ROOT = ROOT.parent.parent

    BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", ROOT))

else:

    ROOT = Path(__file__).resolve().parent

    BUNDLE_DIR = ROOT

def _resolve_web_dir() -> Path:
    candidates = [
        BUNDLE_DIR / "web",
        ROOT / "web",
        BUNDLE_DIR / "_internal" / "web",
        ROOT / "_internal" / "web",
        Path(__file__).resolve().parent / "web",
        Path(__file__).resolve().parent / "_internal" / "web",
        Path(__file__).resolve().parent.parent / "Fome" / "web",
        Path(sys.executable).resolve().parent / "web",
        Path(sys.executable).resolve().parent / "_internal" / "web",
    ]
    for c in candidates:
        if c.is_dir() and (c / "index.html").is_file():
            return c
    return ROOT / "web"

WEB_DIR = _resolve_web_dir()

DATA_DIR = ROOT / "Fome_Web_Data"

BACKUP_DIR = DATA_DIR / "Backup"

OUTPUT_DIR = ROOT / "Fome_Web_Output"

SITES_DIR = ROOT / "Fome_Sites"

DEFAULT_SITE = "TTC15A"

SITE_CODE_RE = re.compile(r"^[A-Z0-9][A-Z0-9_-]{1,30}$")

HOST = os.environ.get("FOME_WEB_HOST", "0.0.0.0")

PORT = int(os.environ.get("PORT", os.environ.get("FOME_WEB_PORT", "8765")))

MAX_UPLOAD_BYTES = 100 * 1024 * 1024

SOURCES = {

    "controlog": {

        "filename": "Controlog WM TTC15A 11-9-2026.xlsx",

        "seed": ROOT / "Controlog WM TTC15A 11-9-2026.xlsx",

        "sheet": "Control Log  TTC15A",

    },

    "rfi_data": {

        "filename": "FORM CHOOSE REPORT Rev 8_ORIGINAL- TTC15A.xlsx",

        "seed": ROOT / "RFI 15A" / "FORM CHOOSE REPORT Rev 8_ORIGINAL- TTC15A.xlsx",

        "sheet": "FORM",

    },

}

def source_filename(key: str, location: str = DEFAULT_SITE) -> str:

    """Return the data filename for a given source key and location.

    For the DEFAULT_SITE we keep the original hardcoded names for backward

    compatibility.  For every other site we generate a site-specific name so

    that files belonging to different locations are never confused.

    """

    if location == DEFAULT_SITE:

        return SOURCES[key]["filename"]

    if key == "controlog":

        return f"Controlog WM {location}.xlsx"

    if key == "rfi_data":

        return f"FORM CHOOSE REPORT - {location}.xlsx"

    return SOURCES[key]["filename"]

RP_RANDOM_MEASUREMENT_TYPES = {"CAI", "DIC", "FTI", "FUR", "HTI", "SPI"}

RANDOM_MEASUREMENT_FORMULA = re.compile(r"\b(?:RAND|RANDBETWEEN)\s*\(", re.IGNORECASE)

FORMULA_CELL_REFERENCE = re.compile(

    r"(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_. ]*))!)?"

    r"(\$?[A-Z]{1,3}\$?\d+)",

    re.IGNORECASE,

)

STATION_CELL_RE = re.compile(

    r"Tr[ạa]m\s+thi\s+c[ôo]ng|Construction\s+station|\bTTC\d+[A-Za-z]?\b",

    re.IGNORECASE,

)

def station_code_for_location(location: str) -> str:

    cleaned = location.strip().upper()

    if cleaned.startswith("TTC"):

        return cleaned

    if cleaned.startswith("T") and len(cleaned) > 1 and cleaned[1].isdigit():

        return f"TTC{cleaned[1:]}"

    if cleaned.isdigit() or re.match(r"^\d+[A-Z]?$", cleaned):

        return f"TTC{cleaned}"

    return cleaned

def is_station_cell(value: object) -> bool:

    if not isinstance(value, str):

        return False

    return bool(STATION_CELL_RE.search(value))

def update_station_text(text: str, station_code: str) -> str:

    if not is_station_cell(text):

        return text

    updated = text

    # 1. Thay mã trạm sau 'Trạm thi công' hoặc 'Tram thi cong' (chấp nhận :, -, /, khoảng trắng...)

    updated = re.sub(

        r"(Tr[ạa]m\s+thi\s+c[ôo]ng\s*[:/-]?\s*)(?:TTC|T)?[A-Za-z0-9_-]+",

        rf"\g<1>{station_code}",

        updated,

        flags=re.IGNORECASE,

    )

    # 2. Thay mã trạm sau 'Construction station'

    updated = re.sub(

        r"(Construction\s+station\s*[:/-]?\s*)(?:TTC|T)?[A-Za-z0-9_-]+",

        rf"\g<1>{station_code}",

        updated,

        flags=re.IGNORECASE,

    )

    # 3. Thay các cụm TTC... còn lại trong ô này nếu chưa khớp ở trên

    updated = re.sub(

        r"\bTTC\d+[A-Za-z]?\b",

        station_code,

        updated,

        flags=re.IGNORECASE,

    )

    return updated

DEFAULT_DRAWING_NO = "SWG-PTSCL18-EPC-70-PL-DAG-0002 Rev 2 SHT 50/94"

DEFAULT_WELD_MAP_NO = "SWG-PTSCL18-CON-70-PL-WMP-0002 Rev 0 SHT 43/94"

def update_drawing_cell_value(val: object, drawing_no: str, weld_map_no: str) -> tuple[str, bool]:

    """Nhận diện và thay thế số bản vẽ và số bản vẽ mối hàn trong ô Excel."""

    if not isinstance(val, str) or not val.strip():

        return str(val or ""), False

    text = val

    # Tự động chuẩn hóa "Bảng vẽ" thành "Bản vẽ"

    text = re.sub(r"B[ảa]ng\s+vẽ", "Bản vẽ", text, flags=re.IGNORECASE)

    # 1. Kiểm tra Bản vẽ mối hàn trước (vì 'Bản vẽ mối hàn' chứa cả từ 'Bản vẽ')

    if re.search(r"Weld\s*map|m[ốo]i\s+h[àa]n", text, re.IGNORECASE):

        if re.search(r"WMP|SHT|:\s*\S", text, re.IGNORECASE):

            new_val = re.sub(r":\s*.*$", f": {weld_map_no}", text)

            if new_val != val:

                return new_val, True

    # 2. Kiểm tra Bản vẽ số / Drawing No.

    elif re.search(r"Drawing\s*No|B[ảa]n\s+vẽ", text, re.IGNORECASE):

        if re.search(r"DAG|SHT|:\s*\S", text, re.IGNORECASE):

            new_val = re.sub(r":\s*.*$", f": {drawing_no}", text)

            if new_val != val:

                return new_val, True

    # 3. Giá trị bản vẽ đứng độc lập

    if re.search(r"WMP-\d+|CON-\d+-PL-WMP", text, re.IGNORECASE):

        if text != weld_map_no:

            return weld_map_no, True

    if re.search(r"DAG-\d+|EPC-\d+-PL-DAG|CON-\d+-PL-DAG", text, re.IGNORECASE):

        if text != drawing_no:

            return drawing_no, True

    if text != val:

        return text, True

    return text, False

def detect_drawings_from_control_log(location: str) -> tuple[str | None, str | None]:

    """Tự động nhận diện số bản vẽ mối hàn (Cột C, E, D) và số bản vẽ (Cột F, H, G) từ Control Log."""

    try:

        p = source_path("controlog", location)

        if not p.is_file():

            return None, None

        wb = load_workbook(p, read_only=True, data_only=True)

        sheet = None

        for name in wb.sheetnames:

            if name.lower().startswith("control log"):

                sheet = wb[name]

                break

        if sheet is None:

            wb.close()

            return None, None

        detected_wmp = None

        detected_dwg = None

        for row in sheet.iter_rows(min_row=4, max_row=35, values_only=True):

            if not row or len(row) < 7:

                continue

            c = str(row[2] or "").strip()

            d = str(row[3] or "").strip()

            e_raw = row[4] if len(row) > 4 else None

            e = str(e_raw).strip() if e_raw is not None and str(e_raw).strip() != "" else ""

            f = str(row[5] if len(row) > 5 else "" or "").strip()

            g = str(row[6] if len(row) > 6 else "" or "").strip()

            h_raw = row[7] if len(row) > 7 else None

            h = str(h_raw).strip() if h_raw is not None and str(h_raw).strip() != "" else ""

            if c and d and not detected_wmp:

                c_clean = re.sub(r"\s+Rev.*$", "", c, flags=re.IGNORECASE).strip()

                w_parts = [c_clean]

                if e and e.upper() not in {"NA", "N/A", "-"}:

                    w_parts.append(f"REV.{e}")

                w_parts.append(f"SHT {d}")

                detected_wmp = " ".join(w_parts)

            if f and g and not detected_dwg:

                f_clean = re.sub(r"\s+Rev.*$", "", f, flags=re.IGNORECASE).strip()

                d_parts = [f_clean]

                if h and h.upper() not in {"NA", "N/A", "-"}:

                    d_parts.append(f"REV.{h}")

                denom = d.split("/", 1)[1] if "/" in d else ""

                g_str = f"{g}/{denom}" if ("/" not in g and denom) else g

                d_parts.append(f"SHT {g_str}")

                detected_dwg = " ".join(d_parts)

            if detected_wmp and detected_dwg:

                break

        wb.close()

        return detected_dwg, detected_wmp

    except Exception:

        return None, None

def detect_weld_map_from_control_log(location: str) -> str | None:

    _, wmp = detect_drawings_from_control_log(location)

    return wmp

def default_engineer_for_location(location: str) -> str:

    loc_clean = re.sub(r"[^A-Za-z0-9]", "", location).upper()

    if "TTC11" in loc_clean or loc_clean == "T11":

        return "Lê Đình Trường"

    return "Nguyễn Đức Sơn"

_SITE_CONFIG_CACHE: dict[str, dict[str, str]] = {}

def get_site_config(location: str) -> dict[str, str]:

    location = validate_site_code(location, require_existing=False)

    cached = _SITE_CONFIG_CACHE.get(location)

    if cached is not None:

        return dict(cached)

    cfg_file = site_data_dir(location) / "site_config.json"

    defaults = {

        "drawing_no": DEFAULT_DRAWING_NO,

        "weld_map_no": DEFAULT_WELD_MAP_NO,

        "engineer_name": default_engineer_for_location(location),

    }

    det_drawing, det_weld = detect_drawings_from_control_log(location)

    if det_drawing:

        defaults["drawing_no"] = det_drawing

    if det_weld:

        defaults["weld_map_no"] = det_weld

    # Giá trị thực tế của TTC15 (theo bản vẽ thi công)

    if location == "TTC15":

        defaults["drawing_no"] = "SWG-PTSCL18-EPC-70-PL-DAG-0002 SHT 17/94"

        if not det_weld:

            defaults["weld_map_no"] = "SWG-PTSCL18-CON-70-PL-WMP-0002 SHT 17/94"

    if cfg_file.is_file():

        try:

            saved = json.loads(cfg_file.read_text(encoding="utf-8"))

            if isinstance(saved, dict):

                if saved.get("drawing_no") and saved["drawing_no"] not in {

                    DEFAULT_DRAWING_NO,

                    "SWG-PTSCL18-CON-70-PL-DAG-0002 Rev 2 SHT 43/94",

                }:

                    defaults["drawing_no"] = str(saved["drawing_no"]).strip()

                elif det_drawing:

                    defaults["drawing_no"] = det_drawing

                if saved.get("weld_map_no") and saved["weld_map_no"] not in {

                    DEFAULT_WELD_MAP_NO,

                    "SWG-PTSCL18-CON-70-PL-WMP-0002 Rev 0 SHT 43/94",

                }:

                    defaults["weld_map_no"] = str(saved["weld_map_no"]).strip()

                elif det_weld:

                    defaults["weld_map_no"] = det_weld

                if saved.get("engineer_name"):

                    defaults["engineer_name"] = str(saved["engineer_name"]).strip()

        except Exception:

            pass

    _SITE_CONFIG_CACHE[location] = dict(defaults)

    return defaults

def save_site_config(

    location: str,

    drawing_no: str,

    weld_map_no: str,

    engineer_name: str | None = None,

) -> dict[str, str]:

    location = validate_site_code(location, require_existing=False)

    cfg_dir = site_data_dir(location)

    cfg_dir.mkdir(parents=True, exist_ok=True)

    cfg_file = cfg_dir / "site_config.json"

    clean_drawing = drawing_no.strip() or DEFAULT_DRAWING_NO

    clean_weld_map = weld_map_no.strip() or DEFAULT_WELD_MAP_NO

    curr_cfg = {}

    if cfg_file.is_file():

        try:

            curr_cfg = json.loads(cfg_file.read_text(encoding="utf-8"))

            if not isinstance(curr_cfg, dict):

                curr_cfg = {}

        except Exception:

            curr_cfg = {}

    if engineer_name is not None and engineer_name.strip():

        clean_engineer = engineer_name.strip()

    elif curr_cfg.get("engineer_name"):

        clean_engineer = str(curr_cfg["engineer_name"]).strip()

    else:

        clean_engineer = default_engineer_for_location(location)

    config = {

        "drawing_no": clean_drawing,

        "weld_map_no": clean_weld_map,

        "engineer_name": clean_engineer,

    }

    cfg_file.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")

    _SITE_CONFIG_CACHE[location] = dict(config)

    return config

class UserError(Exception):

    pass

def ensure_data_files() -> None:

    DATA_DIR.mkdir(exist_ok=True)

    BACKUP_DIR.mkdir(exist_ok=True)

    OUTPUT_DIR.mkdir(exist_ok=True)

    SITES_DIR.mkdir(exist_ok=True)

    for source in SOURCES.values():

        destination = DATA_DIR / source["filename"]

        if source["seed"].is_file():

            if not destination.exists():

                shutil.copy2(source["seed"], destination)

def get_active_site() -> str:

    state_file = DATA_DIR / "app_state.json"

    if state_file.is_file():

        try:

            data = json.loads(state_file.read_text(encoding="utf-8"))

            loc = str(data.get("active_location", "")).strip().upper()

            if loc and (loc == DEFAULT_SITE or loc in list_sites()):

                return loc

        except Exception:

            pass

    sites = list_sites()

    latest_site = DEFAULT_SITE

    latest_mtime = 0.0

    for s in sites:

        for k in ("controlog", "rfi_data"):

            try:

                sp = source_path(k, s)

                if sp.is_file():

                    mt = sp.stat().st_mtime

                    if mt > latest_mtime:

                        latest_mtime = mt

                        latest_site = s

            except Exception:

                pass

        for fam in ("rfi", "rp"):

            try:

                fdir = site_template_dir(s, fam)

                if fdir.is_dir():

                    for tf in fdir.iterdir():

                        if tf.is_file() and tf.suffix.lower() in {".xlsx", ".xlsm"}:

                            mt = tf.stat().st_mtime

                            if mt > latest_mtime:

                                latest_mtime = mt

                                latest_site = s

            except Exception:

                pass

    return latest_site

def set_active_site(location: str) -> None:

    try:

        location = validate_site_code(location, require_existing=False)

        DATA_DIR.mkdir(parents=True, exist_ok=True)

        state_file = DATA_DIR / "app_state.json"

        data = {}

        if state_file.is_file():

            try:

                data = json.loads(state_file.read_text(encoding="utf-8"))

            except Exception:

                data = {}

        data["active_location"] = location

        data["updated_at"] = datetime.now().isoformat()

        state_file.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    except Exception:

        pass

def validate_site_code(value: str, require_existing: bool = False) -> str:

    location = value.strip().upper()

    if not SITE_CODE_RE.fullmatch(location):

        raise UserError("Mã vị trí chỉ được có chữ, số, gạch ngang hoặc gạch dưới.")

    if require_existing and location != DEFAULT_SITE and location not in list_sites():

        raise UserError(f"Chưa tạo hồ sơ cho vị trí {location}.")

    return location

def list_sites() -> list[str]:

    sites = [DEFAULT_SITE]

    if SITES_DIR.is_dir():

        sites.extend(

            child.name

            for child in SITES_DIR.iterdir()

            if child.is_dir() and SITE_CODE_RE.fullmatch(child.name)

        )

    return sorted(set(sites))

def site_root(location: str) -> Path:

    location = validate_site_code(location, require_existing=False)

    return ROOT if location == DEFAULT_SITE else SITES_DIR / location

def site_data_dir(location: str) -> Path:

    location = validate_site_code(location, require_existing=False)

    return DATA_DIR if location == DEFAULT_SITE else site_root(location) / "Fome_Web_Data"

def site_template_dir(location: str, family: str) -> Path:

    if family not in {"rfi", "rp"}:

        raise UserError("Không nhận diện được nhóm biểu mẫu.")

    if location == DEFAULT_SITE:

        folder_name = "RFI 15A" if family == "rfi" else "RP 15A"

        return ROOT / folder_name

    site = SITES_DIR / location

    generic = site / ("RFI" if family == "rfi" else "RP")

    legacy = site / ("RFI 15A" if family == "rfi" else "RP 15A")

    if generic.is_dir():

        return generic

    if legacy.is_dir() and any(legacy.iterdir()):

        return legacy

    return generic

def site_output_dir(location: str) -> Path:

    location = validate_site_code(location, require_existing=False)

    destination = OUTPUT_DIR / location

    destination.mkdir(parents=True, exist_ok=True)

    return destination

def ensure_site(location: str) -> str:

    location = validate_site_code(location, require_existing=False)

    if location != DEFAULT_SITE:

        root = SITES_DIR / location

        data_dir = root / "Fome_Web_Data"

        data_backup = data_dir / "Backup"

        data_backup.mkdir(parents=True, exist_ok=True)

        # Tự động nạp file dữ liệu FORM CHOOSE REPORT cho vị trí mới nếu chưa có

        target_form = data_dir / source_filename("rfi_data", location)

        if not target_form.exists():

            seed_candidates = [

                DATA_DIR / source_filename("rfi_data", DEFAULT_SITE),

                ROOT / "RFI 15A" / "FORM CHOOSE REPORT Rev 8_ORIGINAL- TTC15A.xlsx",

                ROOT / "RFI 15" / "FORM CHOOSE REPORT Rev 8_ORIGINAL- TTC15A.xlsx",

            ]

            if SITES_DIR.is_dir():

                for s in SITES_DIR.iterdir():

                    if s.is_dir() and s.name != location:

                        seed_candidates.append(

                            s / "Fome_Web_Data" / f"FORM CHOOSE REPORT - {s.name}.xlsx"

                        )

                        seed_candidates.append(

                            s / "Fome_Web_Data" / "FORM CHOOSE REPORT Rev 8_ORIGINAL- TTC15A.xlsx"

                        )

            for sc in seed_candidates:

                if sc.is_file():

                    try:

                        shutil.copy2(sc, target_form)

                        break

                    except Exception:

                        pass

        # Tự động nạp file Controlog mẫu cho vị trí mới nếu chưa có

        target_ctrl = data_dir / source_filename("controlog", location)

        if not target_ctrl.exists():

            ctrl_candidates = [

                DATA_DIR / source_filename("controlog", DEFAULT_SITE),

                ROOT / "Controlog WM TTC15A 11-9-2026.xlsx",

            ]

            if SITES_DIR.is_dir():

                for s in SITES_DIR.iterdir():

                    if s.is_dir() and s.name != location:

                        ctrl_candidates.append(

                            s / "Fome_Web_Data" / f"Controlog WM {s.name}.xlsx"

                        )

                        ctrl_candidates.append(

                            s / "Fome_Web_Data" / "Controlog WM TTC15.xlsx"

                        )

            for cc in ctrl_candidates:

                if cc.is_file():

                    try:

                        shutil.copy2(cc, target_ctrl)

                        break

                    except Exception:

                        pass

        # Create template directories and seed them with templates.

        for family in ("rfi", "rp"):

            target_dir = root / ("RFI" if family == "rfi" else "RP")

            target_dir.mkdir(parents=True, exist_ok=True)

            if any(target_dir.iterdir()):

                continue  # Already has templates, don't overwrite.

            suffix = re.sub(r"^TTC", "", location, flags=re.IGNORECASE)

            root_source = ROOT / (

                f"RFI {suffix}" if family == "rfi" else f"RP {suffix}"

            )

            if not root_source.is_dir():

                root_source = site_template_dir(DEFAULT_SITE, family)

            if root_source.is_dir():

                for src_file in root_source.iterdir():

                    if src_file.suffix.lower() in {".xlsx", ".xlsm"}:

                        if src_file.name.startswith("FORM CHOOSE"):

                            continue

                        dest_file = target_dir / src_file.name

                        if not dest_file.exists():

                            try:

                                shutil.copy2(src_file, dest_file)

                                apply_station_code_to_template_file(

                                    dest_file, station_code_for_location(location)

                                )

                                if family == "rp":

                                    cfg = get_site_config(location)

                                    apply_drawings_to_rp_template_file(

                                        dest_file, cfg["drawing_no"], cfg["weld_map_no"]

                                    )

                            except Exception:

                                pass

    return location

def create_site(

    location: str,

    drawing_no: str = "",

    weld_map_no: str = "",

    engineer_name: str = "",

) -> str:

    location = validate_site_code(location, require_existing=False)

    if location == DEFAULT_SITE or location in list_sites():

        raise UserError(f"Hồ sơ {location} đã tồn tại.")

    if drawing_no.strip() or weld_map_no.strip() or engineer_name.strip():

        save_site_config(location, drawing_no, weld_map_no, engineer_name)

    ensure_site(location)

    return location

def source_path(key: str, location: str = DEFAULT_SITE) -> Path:

    if key not in SOURCES:

        raise UserError("Không nhận diện được loại file dữ liệu.")

    location = validate_site_code(location, require_existing=False)

    folder = site_data_dir(location)

    std_name = source_filename(key, location)

    std_path = folder / std_name

    if not folder.is_dir():

        return std_path

    candidates: list[Path] = []

    if std_path.is_file():

        candidates.append(std_path)

    try:

        for item in folder.iterdir():

            if not item.is_file() or item.suffix.lower() not in {".xlsx", ".xlsm"}:

                continue

            if item.name.startswith("~") or item.name.endswith(".upload"):

                continue

            name_lower = item.name.lower()

            if key == "controlog" and ("control" in name_lower or "ctrl" in name_lower):

                candidates.append(item)

            elif key == "rfi_data" and ("form" in name_lower or "rfi" in name_lower):

                candidates.append(item)

    except Exception:

        pass

    if candidates:

        # Prefer files whose stem contains the exact location code but NOT a superset

        # e.g. for location='TTC15', prefer 'FORM - TTC15.xlsx' over 'FORM - TTC15A.xlsx'

        loc_upper = location.upper()

        exact_match = [

            c for c in candidates

            if loc_upper in c.stem.upper()

            and not re.search(re.escape(loc_upper) + r"[A-Z0-9]", c.stem.upper())

        ]

        pool = exact_match if exact_match else candidates

        return max(pool, key=lambda p: p.stat().st_mtime)

    return std_path

def form_sheet_name(workbook) -> str | None:

    def has_rfi_header(sheet) -> bool:

        header = re.sub(

            r"[^a-z0-9]", "", str(sheet.cell(3, 11).value or "").casefold()

        )

        if header in {"rfinumber", "rfino", "sorfi", "rfi"}:

            return True

        for r in range(1, 6):

            for c in range(1, 16):

                val = re.sub(r"[^a-z0-9]", "", str(sheet.cell(r, c).value or "").casefold())

                if val in {"rfinumber", "rfino", "sorfi"}:

                    return True

        return False

    if "FORM" in workbook.sheetnames:

        return "FORM"

    for name in workbook.sheetnames:

        if "form" in name.casefold():

            return name

    for sheet in workbook.worksheets:

        if has_rfi_header(sheet):

            return sheet.title

    return workbook.sheetnames[0] if workbook.sheetnames else None

def load_records(

    location: str = DEFAULT_SITE,

) -> tuple[list[dict[str, str]], dict[str, dict[int, object]]]:

    location = validate_site_code(location, require_existing=False)

    path = source_path("rfi_data", location)

    if not path.is_file():

        raise UserError(

            f"Vị trí {location} chưa có file nguồn {source_filename('rfi_data', location)}."

        )

    try:

        workbook = load_workbook(path, read_only=True, data_only=True)

    except Exception as exc:

        raise UserError(f"Không đọc được file danh sách RFI: {exc}") from exc

    sheet_name = form_sheet_name(workbook)

    if sheet_name is None:

        workbook.close()

        raise UserError(

            "Không tìm thấy sheet FORM. Cần sheet FORM hoặc sheet có tiêu đề "

            "'RFI Number' tại ô K3."

        )

    sheet = workbook[sheet_name]

    normalized_header = re.sub(

        r"[^a-z0-9]", "", str(sheet.cell(3, 11).value or "").casefold()

    )

    if normalized_header not in {"rfinumber", "rfino", "sorfi", "rfi"}:

        found_hdr = False

        for r in range(1, 6):

            for c in range(1, 16):

                val = re.sub(r"[^a-z0-9]", "", str(sheet.cell(r, c).value or "").casefold())

                if val in {"rfinumber", "rfino", "sorfi"}:

                    found_hdr = True

                    break

            if found_hdr:

                break

        if not found_hdr and "form" not in sheet_name.casefold():

            workbook.close()

            raise UserError(

                f"Sheet {sheet_name} cần tiêu đề 'RFI Number' tại ô K3."

            )

    records: list[dict[str, str]] = []

    rows_by_id: dict[str, dict[int, object]] = {}

    duplicate_ids: list[str] = []

    value_width = max(sheet.max_column, 30)

    for row in sheet.iter_rows(

        min_row=4, max_row=sheet.max_row, min_col=1, max_col=value_width, values_only=True

    ):

        rfi_number = row[10]

        if rfi_number is None or not str(rfi_number).strip():

            continue

        rfi_id = str(rfi_number).strip()

        if rfi_id in rows_by_id:

            # Keep first occurrence, skip duplicate — record warning instead of crashing

            if rfi_id not in duplicate_ids:

                duplicate_ids.append(rfi_id)

            continue

        values = {column: value for column, value in enumerate(row, start=1)}

        rows_by_id[rfi_id] = values

        report_type = get_rfi_type(rfi_id)

        desc_text = cell_text(values.get(12))

        is_rp_supp = bool(

            report_type

            and find_rp_template(report_type, location)

        )

        rp_note_str = (

            "Giữ nguyên công thức đo của biểu mẫu."

            if report_type in RP_RANDOM_MEASUREMENT_TYPES

            else ""

        )

        if report_type == "MRIR":

            if "welding" not in desc_text.lower():

                is_rp_supp = False

                rp_note_str = "Vật tư ống (Pipe) - Không cần xuất RP."

        records.append(

            {

                "id": rfi_id,

                "type": report_type or "",

                "description": desc_text,

                "report": cell_text(values.get(13)),

                "date": cell_text(values.get(14)),

                "supported": bool(report_type and find_template(report_type, location)),

                "rp_supported": is_rp_supp,

                "rp_note": rp_note_str,

            }

        )

    workbook.close()

    return records, rows_by_id, duplicate_ids

def cell_text(value: object) -> str:

    if value is None:

        return ""

    if isinstance(value, (datetime, date)):

        return value.strftime("%Y-%m-%d")

    return str(value).strip()

def get_rfi_type(rfi_id: str) -> str | None:

    match = re.search(r"-RFI-([A-Z0-9]+)-", rfi_id, flags=re.IGNORECASE)

    if match:

        return match.group(1).upper()

    match = re.search(r"-([A-Z]{3,4})-\d{4}$", rfi_id, flags=re.IGNORECASE)

    if match:

        return match.group(1).upper()

    return None

def find_template(report_type: str, location: str = DEFAULT_SITE) -> Path | None:

    folder = site_template_dir(location, "rfi")

    if not folder.is_dir():

        return None

    matches: list[Path] = []

    for candidate in folder.iterdir():

        if candidate.suffix.lower() not in {".xlsx", ".xlsm"}:

            continue

        if candidate.name.startswith("FORM CHOOSE"):

            continue

        match = re.search(r"RFI[\s-]*([A-Z0-9]+)", candidate.stem, flags=re.IGNORECASE)

        if match and match.group(1).upper() == report_type:

            matches.append(candidate)

    if matches:

        return max(matches, key=lambda p: p.stat().st_mtime)

    return None

def find_rp_template(report_type: str, location: str = DEFAULT_SITE) -> Path | None:

    folder = site_template_dir(location, "rp")

    if not folder.is_dir():

        return None

    matches: list[Path] = []

    for candidate in folder.iterdir():

        if candidate.suffix.lower() not in {".xlsx", ".xlsm"}:

            continue

        match = re.search(r"\bRP\s*-\s*([A-Za-z]+)(?![A-Za-z0-9])", candidate.stem, flags=re.IGNORECASE)

        if match and match.group(1).upper() == report_type:

            matches.append(candidate)

    if matches:

        return max(matches, key=lambda p: p.stat().st_mtime)

    return None

def excel_column_index(column: str) -> int:

    try:

        return column_index_from_string(column.upper())

    except ValueError as exc:

        raise UserError(f"Công thức biểu mẫu chứa cột không hợp lệ: {column}") from exc

LOOKUP_PATTERN = re.compile(

    r"VLOOKUP\(\s*([^,]+?)\s*,\s*\[\d+\]FORM!"

    r"([A-Z]{1,3})\$?(\d+):([A-Z]{1,3})\$?\d+\s*,\s*(\d+)\s*,\s*0\s*\)",

    flags=re.IGNORECASE,

)

def evaluate_form_formula(

    formula: str,

    selected_id: str,

    form_rows: dict[str, dict[int, object]],

) -> object:

    lookups = list(LOOKUP_PATTERN.finditer(formula))

    if not lookups:

        raise UserError(f"Không hỗ trợ công thức trong biểu mẫu: {formula}")

    output_values: list[object] = []

    for match in lookups:

        lookup_value = match.group(1).strip()

        if lookup_value.upper() not in {"A6", "$A$6"}:

            raise UserError(f"Không hỗ trợ ô dò tìm trong công thức: {formula}")

        start_column = excel_column_index(match.group(2))

        result_column = start_column + int(match.group(5)) - 1

        source_row = form_rows.get(selected_id)

        if source_row is None:

            raise UserError(f"Không tìm thấy số RFI {selected_id} trong file dữ liệu.")

        output_values.append(source_row.get(result_column))

    if len(lookups) == 1 and formula.strip().startswith("=VLOOKUP("):

        return output_values[0]

    expression = formula.lstrip("=").strip()

    tokens = iter(output_values)

    expression = LOOKUP_PATTERN.sub(lambda _: f"\x00{next(tokens)}\x00", expression)

    pieces = re.split(r'&(?=(?:[^"]*"[^"]*")*[^"]*$)', expression)

    joined: list[str] = []

    for piece in pieces:

        token = piece.strip()

        if token.startswith("\x00") and token.endswith("\x00"):

            joined.append(cell_text(token[1:-1]))

        elif len(token) >= 2 and token.startswith('"') and token.endswith('"'):

            joined.append(token[1:-1].replace('""', '"'))

        elif token:

            raise UserError(f"Không hỗ trợ phép tính trong công thức: {formula}")

    return "".join(joined)

def cell_xml_value(value: object) -> tuple[str, str]:

    if value is None:

        return "", ""

    if isinstance(value, bool):

        return ' t="b"', "1" if value else "0"

    if isinstance(value, (datetime, date)):

        if isinstance(value, datetime):

            value = value.replace(tzinfo=None)

        serial = (value - datetime(1899, 12, 30)).total_seconds() / 86400

        return "", format(serial, ".15g")

    if isinstance(value, (int, float)):

        return "", format(value, ".15g")

    return ' t="str"', html.escape(str(value), quote=False)

def replace_cell_content(xml_bytes: bytes, reference: str, value: object) -> bytes:

    reference_pattern = re.escape(reference)

    cell_pattern = re.compile(

        rb"<c\b(?=[^>]*\br=(['\"])" + reference_pattern.encode("ascii") +

        rb"\1)[^>]*(?:/>|>.*?</c>)",

        flags=re.DOTALL,

    )

    match = cell_pattern.search(xml_bytes)

    if match is None:

        raise UserError(f"Không tìm thấy ô {reference} trong file mẫu.")

    cell = match.group(0)

    cell = re.sub(rb'\s+t="[^"]*"', b"", cell, count=1)

    cell = re.sub(rb"\s+t='[^']*'", b"", cell, count=1)

    cell = re.sub(rb"<f(?:\s[^>]*)?/?>.*?</f\s*>", b"", cell, flags=re.DOTALL)

    cell = re.sub(rb"<f(?:\s[^>]*)?/>", b"", cell)

    cell = re.sub(rb"<(?:v|is)(?:\s[^>]*)?/?>.*?</(?:v|is)\s*>", b"", cell, flags=re.DOTALL)

    cell = re.sub(rb"<(?:v|is)(?:\s[^>]*)?/>", b"", cell)

    if cell.endswith(b"/>"):

        cell = cell[:-2] + b"></c>"

    if reference == "A6" or (isinstance(value, str) and ("\n" in value or is_station_cell(value) or "WMP" in value or "DAG" in value or "SHT" in value)):

        type_attribute = b' t="inlineStr"'

        cell = re.sub(rb"(<c\b)", rb"\1" + type_attribute, cell, count=1)

        text = html.escape(str(value), quote=False).encode("utf-8")

        payload = b'<is><t xml:space="preserve">' + text + b"</t></is>"

        cell = re.sub(rb"(</c>)$", payload + rb"\1", cell)

    else:

        type_attribute, text_value = cell_xml_value(value)

        if type_attribute:

            cell = re.sub(rb"(<c\b)", rb"\1" + type_attribute.encode("ascii"), cell, count=1)

        payload = b"<v>" + text_value.encode("utf-8") + b"</v>"

        cell = re.sub(rb"(</c>)$", payload + rb"\1", cell)

    return xml_bytes[: match.start()] + cell + xml_bytes[match.end() :]

def _replace_cell_match(match: re.Match[bytes], reference: str, value: object) -> bytes:

    cell = match.group(0)

    reference_match = re.search(rb'\br=(["\'])([A-Z]{1,3}\d+)\1', cell)

    if reference_match is None or reference_match.group(2).decode("ascii") != reference:

        return cell

    cell = re.sub(rb'\s+t="[^"]*"', b"", cell, count=1)

    cell = re.sub(rb"\s+t='[^']*'", b"", cell, count=1)

    cell = re.sub(rb"<f(?:\s[^>]*)?/?>.*?</f\s*>", b"", cell, flags=re.DOTALL)

    cell = re.sub(rb"<f(?:\s[^>]*)?/>", b"", cell)

    cell = re.sub(rb"<(?:v|is)(?:\s[^>]*)?/?>.*?</(?:v|is)\s*>", b"", cell, flags=re.DOTALL)

    cell = re.sub(rb"<(?:v|is)(?:\s[^>]*)?/>", b"", cell)

    if cell.endswith(b"/>"):

        cell = cell[:-2] + b"></c>"

    if isinstance(value, str) and ("\n" in value or is_station_cell(value) or "WMP" in value or "DAG" in value or "SHT" in value):

        type_attribute = b' t="inlineStr"'

        cell = re.sub(rb"(<c\b)", rb"\1" + type_attribute, cell, count=1)

        text = html.escape(str(value), quote=False).encode("utf-8")

        payload = b'<is><t xml:space="preserve">' + text + b"</t></is>"

    elif value is None or value == "":

        payload = b"<v></v>"

    else:

        type_attribute, text_value = cell_xml_value(value)

        if type_attribute:

            cell = re.sub(rb"(<c\b)", rb"\1" + type_attribute.encode("ascii"), cell, count=1)

        payload = b"<v>" + text_value.encode("utf-8") + b"</v>"

    return re.sub(rb"(</c>)$", payload + rb"\1", cell)

def replace_cells_content(xml_bytes: bytes, replacements: dict[str, object]) -> bytes:

    if not replacements:

        return xml_bytes

    cell_pattern = re.compile(rb"<c\b[^>]*?(?:/>|>.*?</c>)", flags=re.DOTALL)

    def replace_match(match: re.Match[bytes]) -> bytes:

        reference_match = re.search(

            rb'\br=(["\'])([A-Z]{1,3}\d+)\1', match.group(0)

        )

        if reference_match is None:

            return match.group(0)

        reference = reference_match.group(2).decode("ascii")

        if reference not in replacements:

            return match.group(0)

        return _replace_cell_match(match, reference, replacements[reference])

    return cell_pattern.sub(replace_match, xml_bytes)

def clean_workbook_xml(data: bytes) -> bytes:

    data = re.sub(

        rb'<(?:\w+:)?externalReferences\b[^>]*>.*?</(?:\w+:)?externalReferences\s*>',

        b"",

        data,

        flags=re.DOTALL,

    )

    data = re.sub(rb'<(?:\w+:)?externalReferences\b[^>]*/>', b"", data)

    valid_sheets = set(re.findall(r'<sheet\b[^>]*name="([^"]+)"', data.decode("utf-8", errors="ignore")))

    def filter_dn(m):

        content = m.group(0)

        text = m.group(2).decode("utf-8", errors="ignore")

        if re.search(r"\[\d+\]", text):

            return b""

        if "#REF!" in text or "#N/A" in text:

            return b""

        sheet_refs = [s.strip("'") for s in re.findall(r"'?([^'\"!]+)'?!", text)]

        for s in sheet_refs:

            if s not in valid_sheets:

                return b""

        if ":\\" in text or "장비" in text:

            return b""

        return content

    data = re.sub(rb'<definedName\b([^>]*)>(.*?)</definedName>', filter_dn, data, flags=re.DOTALL)

    data = re.sub(rb'<definedNames\b[^>]*>\s*</definedNames>', b"", data)

    return data

def clean_styles_xml(data: bytes, preserve_red: bool = False) -> bytes:

    """Đảm bảo các style trong bảng có căn giữa ngang và dọc (horizontal/vertical center) và màu đen mặc định."""

    def fix_align(m: re.Match[bytes]) -> bytes:

        tag = m.group(0)

        if b'horizontal=' not in tag:

            return re.sub(rb"(<alignment\b)", rb'\1 horizontal="center"', tag, count=1)

        return tag

    data = re.sub(rb'<alignment\b[^>]*vertical="center"[^>]*wrapText="1"[^>]*/>', fix_align, data)

    data = re.sub(rb'<alignment\b[^>]*wrapText="1"[^>]*vertical="center"[^>]*/>', fix_align, data)

    if not preserve_red:

        # Chuyển các mã màu chữ đỏ hoặc khác về màu đen mặc định FF000000

        data = re.sub(rb'<color\b[^>]*rgb="(?:FFFF0000|FF0000|ff0000|ffff0000)"[^>]*/>', rb'<color rgb="FF000000"/>', data)

    return data

def worksheet_paths(archive: zipfile.ZipFile) -> dict[str, str]:

    try:

        workbook_xml = ElementTree.fromstring(archive.read("xl/workbook.xml"))

        relationships_xml = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))

    except (KeyError, ElementTree.ParseError) as exc:

        raise UserError("Biểu mẫu Excel bị lỗi cấu trúc.") from exc

    relationship_tag = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

    relationship_map = {

        item.attrib["Id"]: item.attrib["Target"]

        for item in relationships_xml

        if item.attrib.get("Id") and item.attrib.get("Target")

    }

    result: dict[str, str] = {}

    for sheet in workbook_xml.findall(".//{*}sheet"):

        name = sheet.attrib.get("name")

        relationship_id = sheet.attrib.get(relationship_tag)

        if name and relationship_id in relationship_map:

            target = relationship_map[relationship_id].replace("\\", "/")

            result[name] = (

                target.lstrip("/")

                if target.startswith("/")

                else posixpath.normpath(posixpath.join("xl", target))

            )

    return result

def sheet_drawing_paths(archive: zipfile.ZipFile) -> dict[str, str]:

    ws_paths = worksheet_paths(archive)

    sheet_to_drawing: dict[str, str] = {}

    for sheet_name, spath in ws_paths.items():

        rels_path = posixpath.normpath(

            posixpath.join(

                posixpath.dirname(spath),

                "_rels",

                posixpath.basename(spath) + ".rels",

            )

        )

        if rels_path in archive.namelist():

            try:

                s_rels = ElementTree.fromstring(archive.read(rels_path))

                for r in s_rels:

                    if "drawing" in r.attrib.get("Type", ""):

                        target = r.attrib.get("Target", "").replace("\\", "/")

                        drawing_path = posixpath.normpath(

                            posixpath.join(posixpath.dirname(spath), target)

                        )

                        sheet_to_drawing[sheet_name] = drawing_path.lstrip("/")

            except Exception:

                pass

    return sheet_to_drawing

HORIZONTAL_RP_GEOMETRY = {

    "CAI": {
        "0001": {
            "joint_row": 14,
            "data_cols": [9, 10, 11, 12, 13, 14],  # 1-based I..N
            "from_row": 12,  # 0-based row 13
            "to_row": 17,    # 0-based row 18
            "base_col": 8,   # 0-based col I
            "end_col": 13,   # 0-based col N
            "max_joints": 6,
        },
        "7.RP-CAI": {
            "joint_row": 14,
            "data_cols": [9, 10, 11, 12, 13, 14],  # 1-based I..N
            "from_row": 12,  # 0-based row 13
            "to_row": 17,    # 0-based row 18
            "base_col": 8,   # 0-based col I
            "end_col": 13,   # 0-based col N
            "max_joints": 6,
        },
        "RP-CAI": {
            "joint_row": 14,
            "data_cols": [9, 10, 11, 12, 13, 14],  # 1-based I..N
            "from_row": 12,  # 0-based row 13
            "to_row": 17,    # 0-based row 18
            "base_col": 8,   # 0-based col I
            "end_col": 13,   # 0-based col N
            "max_joints": 6,
        },
    },

    "SPI": {

        "RP-SPI": {

            "joint_row": 11,

            "data_cols": [8, 9, 10, 11, 12, 13],   # 1-based H..M

            "from_row": 9,   # 0-based row 10

            "to_row": 19,    # 0-based row 20

            "base_col": 7,   # 0-based col H

            "end_col": 12,   # 0-based col M

            "max_joints": 6,

        },

        "0001-1": {

            "joint_row": 11,

            "data_cols": [9, 10, 11, 12, 13, 14],  # 1-based I..N

            "from_row": 9,   # 0-based row 10

            "to_row": 16,    # 0-based row 17

            "base_col": 8,   # 0-based col I

            "end_col": 13,   # 0-based col N

            "max_joints": 6,

        },

    },

    "FTI": {

        "RP-FTI": {

            "joint_row": 10,

            "data_cols": [10, 11, 12, 13, 14, 15],  # 1-based J..O

            "from_row": 8,   # 0-based row 9

            "to_row": 15,    # 0-based row 16

            "base_col": 9,   # 0-based col J

            "end_col": 14,   # 0-based col O

            "max_joints": 6,

        },

        "0001-1": {

            "joint_row": 14,

            "data_cols": [10, 11, 12, 13, 14, 15],  # 1-based J..O

            "from_row": 13,  # 0-based row 14 (below material parameters)

            "to_row": 19,    # 0-based row 20

            "base_col": 9,   # 0-based col J

            "end_col": 14,   # 0-based col O

            "max_joints": 6,

        },

    },

    "HTI": {

        "RP-HTI": {

            "joint_row": 13,

            "data_cols": [10, 11, 12, 13, 14, 15, 16],  # 1-based J..P

            "from_row": 11,  # 0-based row 12 (Time)

            "to_row": 19,    # 0-based row 20 (Sound test)

            "base_col": 9,   # 0-based col J

            "end_col": 15,   # 0-based col P

            "max_joints": 7,

        },

    },

}

def apply_horizontal_diagonal_line(

    drawing_xml: bytes,

    report_type: str,

    sheet_name: str,

    joint_count: int,

    page_len: int = 31,

    num_pages: int = 1,

) -> bytes:

    configs = HORIZONTAL_RP_GEOMETRY.get(report_type)

    if not configs:

        return drawing_xml

    cfg = configs.get(sheet_name)
    if not cfg:
        for k, v in configs.items():
            if k.upper() in sheet_name.upper() or sheet_name.upper() in k.upper():
                cfg = v
                break
    if not cfg:
        if report_type == "MTR":
            for p in range(1, max(1, n_pages) + 1):
                off = (p - 1) * p_len
                # K3 (Date/ Ngày:): căn lề trái
                c_date = ws.cell(row=3 + off, column=11)
                if not isinstance(c_date, MergedCell):
                    c_date.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                # A7 (Vị trí/ Location:): căn lề trái
                c_loc = ws.cell(row=7 + off, column=1)
                if not isinstance(c_loc, MergedCell):
                    f_loc_b = InlineFont(rFont="Times New Roman", sz=10, b=True, i=False)
                    f_loc_i = InlineFont(rFont="Times New Roman", sz=10, b=False, i=True)
                    c_loc.value = CellRichText(TextBlock(f_loc_b, "Vị trí/\n"), TextBlock(f_loc_i, "Location:"))
                    c_loc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                # D7 (Hạng mục / Item:): căn lề trái
                c_item = ws.cell(row=7 + off, column=4)
                if not isinstance(c_item, MergedCell):
                    c_item.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                # E9 (Ống số/ Pipe No.): chữ thường không in đậm
                c_pipe = ws.cell(row=9 + off, column=5)
                if not isinstance(c_pipe, MergedCell):
                    c_pipe.font = Font(name="Times New Roman", size=10.0, bold=False, italic=False, color="000000")

                # G9 (Số nhận dạng/ Serial No.): chữ thường không in đậm
                c_ser = ws.cell(row=9 + off, column=7)
                if not isinstance(c_ser, MergedCell):
                    c_ser.font = Font(name="Times New Roman", size=9.0, bold=False, italic=False, color="000000")
        if report_type == "CAI":
            cfg = configs.get("7.RP-CAI") or configs.get("0001") or next(iter(configs.values()), None)

        elif report_type == "HTI":

            cfg = configs.get("RP-HTI")

        elif report_type == "SPI":

            if "SPI" in sheet_name.upper():

                cfg = configs.get("RP-SPI")

            else:

                cfg = configs.get("0001-1")

        elif report_type == "FTI":

            if "FTI" in sheet_name.upper():

                cfg = configs.get("RP-FTI")

            else:

                cfg = configs.get("0001-1")

    if not cfg:

        return drawing_xml

    ElementTree.register_namespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")

    ElementTree.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")

    ElementTree.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    ElementTree.register_namespace("a14", "http://schemas.microsoft.com/office/drawing/2010/main")

    ElementTree.register_namespace("a16", "http://schemas.microsoft.com/office/drawing/2014/main")

    try:

        root = ElementTree.fromstring(drawing_xml)

    except Exception:

        return drawing_xml

    xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"

    # Remove ONLY connectors / line anchors, NEVER touch pictures or logo groups!

    for child in list(root):

        is_connector = False

        if child.find(f".//{{{xdr_ns}}}cxnSp") is not None:

            is_connector = True

        else:

            geom = child.find(f".//{{{a_ns}}}prstGeom")

            if geom is not None and geom.attrib.get("prst") == "line":

                is_connector = True

            elif child.find(f".//{{{xdr_ns}}}nvCxnSpPr") is not None:

                is_connector = True

        if is_connector:

            root.remove(child)

    # For HTI, user requested to remove the blank cell locking diagonal slash

    if report_type == "HTI":

        return ElementTree.tostring(root, encoding="utf-8")

    max_joints = cfg.get("max_joints", 6)

    base_col = cfg["base_col"]

    end_col = cfg["end_col"]

    from_row = cfg["from_row"]

    to_row = cfg["to_row"]

    max_id = 100

    for cnv in root.iter(f"{{{xdr_ns}}}cNvPr"):

        if "id" in cnv.attrib:

            try:

                max_id = max(max_id, int(cnv.attrib["id"]))

            except ValueError:

                pass

    if num_pages > 1:
        base_anchors = []
        for a in list(root.findall(f"{{{xdr_ns}}}twoCellAnchor")):
            r_from = a.find(f"{{{xdr_ns}}}from/{{{xdr_ns}}}row")
            if r_from is not None and int(r_from.text) < page_len:
                base_anchors.append(a)
        for p_idx in range(2, num_pages + 1):
            offset_val = (p_idx - 1) * page_len
            for a in base_anchors:
                clone = copy.deepcopy(a)
                rf = clone.find(f"{{{xdr_ns}}}from/{{{xdr_ns}}}row")
                rt = clone.find(f"{{{xdr_ns}}}to/{{{xdr_ns}}}row")
                if rf is not None:
                    rf.text = str(int(rf.text) + offset_val)
                if rt is not None:
                    rt.text = str(int(rt.text) + offset_val)
                max_id += 1
                cnvpr = clone.find(f".//{{{xdr_ns}}}cNvPr")
                if cnvpr is not None:
                    cnvpr.attrib["id"] = str(max_id)
                root.append(clone)

    for p in range(1, num_pages + 1):

        p_offset = (p - 1) * page_len

        p_count = max(0, min(max_joints, joint_count - (p - 1) * max_joints))

        if p_count < max_joints:

            start_col = base_col + p_count

            if start_col <= end_col:

                max_id += 1

                anchor = ElementTree.SubElement(root, f"{{{xdr_ns}}}twoCellAnchor")

                frm = ElementTree.SubElement(anchor, f"{{{xdr_ns}}}from")

                c1 = ElementTree.SubElement(frm, f"{{{xdr_ns}}}col"); c1.text = str(start_col)

                co1 = ElementTree.SubElement(frm, f"{{{xdr_ns}}}colOff"); co1.text = "0"

                r1 = ElementTree.SubElement(frm, f"{{{xdr_ns}}}row"); r1.text = str(from_row + p_offset)

                ro1 = ElementTree.SubElement(frm, f"{{{xdr_ns}}}rowOff"); ro1.text = "0"

                to_el = ElementTree.SubElement(anchor, f"{{{xdr_ns}}}to")

                c2 = ElementTree.SubElement(to_el, f"{{{xdr_ns}}}col"); c2.text = str(end_col + 1)

                co2 = ElementTree.SubElement(to_el, f"{{{xdr_ns}}}colOff"); co2.text = "0"

                r2 = ElementTree.SubElement(to_el, f"{{{xdr_ns}}}row"); r2.text = str(to_row + 1 + p_offset)

                ro2 = ElementTree.SubElement(to_el, f"{{{xdr_ns}}}rowOff"); ro2.text = "0"

                cxn = ElementTree.SubElement(anchor, f"{{{xdr_ns}}}cxnSp", attrib={"macro": ""})

                nv = ElementTree.SubElement(cxn, f"{{{xdr_ns}}}nvCxnSpPr")

                ElementTree.SubElement(nv, f"{{{xdr_ns}}}cNvPr", attrib={"id": str(max_id), "name": f"Straight Connector {max_id}"})

                ElementTree.SubElement(nv, f"{{{xdr_ns}}}cNvCxnSpPr")

                spPr = ElementTree.SubElement(cxn, f"{{{xdr_ns}}}spPr")

                xfrm = ElementTree.SubElement(spPr, f"{{{a_ns}}}xfrm")

                ElementTree.SubElement(xfrm, f"{{{a_ns}}}off", attrib={"x": "0", "y": "0"})

                ElementTree.SubElement(xfrm, f"{{{a_ns}}}ext", attrib={"cx": "100000", "cy": "100000"})

                ElementTree.SubElement(spPr, f"{{{a_ns}}}prstGeom", attrib={"prst": "line"})

                ln = ElementTree.SubElement(spPr, f"{{{a_ns}}}ln", attrib={"w": "12700"})

                sf = ElementTree.SubElement(ln, f"{{{a_ns}}}solidFill")

                ElementTree.SubElement(sf, f"{{{a_ns}}}srgbClr", attrib={"val": "000000"})

                style = ElementTree.SubElement(cxn, f"{{{xdr_ns}}}style")

                lnRef = ElementTree.SubElement(style, f"{{{a_ns}}}lnRef", attrib={"idx": "1"})

                ElementTree.SubElement(lnRef, f"{{{a_ns}}}schemeClr", attrib={"val": "dk1"})

                fillRef = ElementTree.SubElement(style, f"{{{a_ns}}}fillRef", attrib={"idx": "0"})

                ElementTree.SubElement(fillRef, f"{{{a_ns}}}schemeClr", attrib={"val": "dk1"})

                effectRef = ElementTree.SubElement(style, f"{{{a_ns}}}effectRef", attrib={"idx": "0"})

                ElementTree.SubElement(effectRef, f"{{{a_ns}}}schemeClr", attrib={"val": "dk1"})

                fontRef = ElementTree.SubElement(style, f"{{{a_ns}}}fontRef", attrib={"idx": "minor"})

                ElementTree.SubElement(fontRef, f"{{{a_ns}}}schemeClr", attrib={"val": "tx1"})

                ElementTree.SubElement(anchor, f"{{{xdr_ns}}}clientData")

    return ElementTree.tostring(root, encoding="utf-8")

def replace_area_with_location(ws) -> None:

    """Thay thế 'Khu vực/Area:' bằng 'Địa điểm/ Location:' trên tất cả các ô tiêu đề."""

    if ws.title.casefold().startswith("controlog"):

        return

    limit_r = min(25, ws.max_row)

    limit_c = min(32, ws.max_column)

    for r in range(1, limit_r + 1):

        for c in range(1, limit_c + 1):

            cell = ws.cell(row=r, column=c)

            if isinstance(cell, MergedCell):

                continue

            val = cell.value

            if isinstance(val, str) and ("khu vực" in val.lower() or "area" in val.lower()):

                new_val = re.sub(

                    r"Khu\s*v[ựu]c\s*/\s*Area\s*:\s*",

                    "Địa điểm/ Location: ",

                    val,

                    flags=re.IGNORECASE,

                )

                if new_val == val:

                    new_val = re.sub(

                        r"Khu\s*v[ựu]c\s*/\s*Area\b",

                        "Địa điểm/ Location",

                        val,

                        flags=re.IGNORECASE,

                    )

                if new_val != val:

                    cell.value = new_val

def format_rp_multiline_slash(val: object) -> str | None:

    """Định dạng xuống dòng cho các nội dung phân biệt bởi dấu '/' (ví dụ trong MTR).

    Không áp dụng cho 'N/A' (giữ nguyên không xuống dòng).

    """

    if not isinstance(val, str) or "/" not in val or "\n" in val:

        return None

    s = val.strip()

    if s.upper() in {"N/A", "NA", "N / A", "N/ A", "N /A"}:

        return None

    if re.fullmatch(r"N\s*/\s*A", s, re.IGNORECASE):

        return None

    if re.match(r"^\d{1,2}/\d{1,2}/\d{2,4}$", s):

        return None

    if re.match(r"^\d+/\d+$", s):

        return None

    if re.search(r"SHT\s*\d+/\d+", s, re.IGNORECASE):

        return None

    if re.search(r"Địa\s*điểm|Location|Khu\s*vực|Area|Hạng\s*mục|Item|Bản\s*vẽ|Drawing", s, re.IGNORECASE):

        return None

    return re.sub(r"\s*/\s*", "/\n", s)

def adjust_drawing_lines(

    drawing_xml: bytes,

    sheet_name: str,

    workbook,

    replacements: dict[str, dict[str, object]],

) -> bytes:

    sheet_rep = replacements.get(sheet_name, {})

    ws = workbook[sheet_name] if (workbook and sheet_name in workbook.sheetnames) else None

    def get_cell_val(r: int, c: int):

        coord = f"{get_column_letter(c)}{r}"

        if coord in sheet_rep:

            return sheet_rep[coord]

        if ws is not None:

            return ws.cell(r, c).value

        return None

    xml_str = drawing_xml.decode("utf-8")

    anchor_pattern = re.compile(

        r"<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>", flags=re.DOTALL

    )

    def process_anchor(match: re.Match) -> str:

        anchor_xml = match.group(0)

        if (

            'prst="line"' not in anchor_xml

            and "Connector" not in anchor_xml

            and "Thẳng" not in anchor_xml

        ):

            return anchor_xml

        m_from = re.search(

            r"<xdr:from>.*?<xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>.*?</xdr:from>",

            anchor_xml,

            flags=re.DOTALL,

        )

        m_to = re.search(

            r"<xdr:to>.*?<xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>.*?</xdr:to>",

            anchor_xml,

            flags=re.DOTALL,

        )

        if not m_from or not m_to:

            return anchor_xml

        from_col = int(m_from.group(1))

        from_row = int(m_from.group(2))

        to_col = int(m_to.group(1))

        to_row = int(m_to.group(2))

        # Check if table is vertical (rows) or horizontal (cols)

        if abs(to_col - from_col) >= 5 and to_row > from_row:

            table_end_row = to_row + 1  # 1-based

            last_data_row = 0

            for r in range(1, table_end_row + 1):

                val_b = get_cell_val(r, 2)

                val_a = get_cell_val(r, 1)

                is_data = False

                if val_b is not None and str(val_b).strip() not in {"", "/"}:

                    s_b = str(val_b).strip().lower()

                    if (

                        "số mối hàn" not in s_b

                        and "joint" not in s_b

                        and "instrument" not in s_b

                    ):

                        is_data = True

                elif val_a is not None and str(val_a).strip() not in {"", "/"}:

                    s_a = str(val_a).strip()

                    if s_a.isdigit() and int(s_a) > 0:

                        is_data = True

                if is_data:

                    last_data_row = r

            if last_data_row == 0:

                new_from_row = from_row

            else:

                first_empty_row = last_data_row + 1

                new_from_row = first_empty_row - 1  # 0-based

            if new_from_row > to_row:

                # All rows are filled with data; remove the line shape

                return ""

            def rep_from(m_f):

                return re.sub(

                    r"(<xdr:row>)(\d+)(</xdr:row>)",

                    rf"\g<1>{new_from_row}\g<3>",

                    m_f.group(0),

                )

            return re.sub(

                r"<xdr:from>.*?</xdr:from>", rep_from, anchor_xml, flags=re.DOTALL

            )

        elif to_col > from_col:

            # Horizontal table (e.g. CAI, FTI)

            table_end_col = to_col + 1

            last_data_col = 0

            for c in range(from_col + 1, table_end_col + 1):

                has_col_data = False

                for r in [10, 14]:

                    val = get_cell_val(r, c)

                    if val is not None and str(val).strip() not in {"", "/"}:

                        s_val = str(val).strip().lower()

                        if "mối" not in s_val and "joint" not in s_val:

                            has_col_data = True

                            break

                if has_col_data:

                    last_data_col = c

            if last_data_col == 0:

                new_from_col = from_col

            else:

                new_from_col = (last_data_col + 1) - 1

            if new_from_col > to_col:

                return ""

            def rep_from_col(m_f):

                return re.sub(

                    r"(<xdr:col>)(\d+)(</xdr:col>)",

                    rf"\g<1>{new_from_col}\g<3>",

                    m_f.group(0),

                )

            return re.sub(

                r"<xdr:from>.*?</xdr:from>", rep_from_col, anchor_xml, flags=re.DOTALL

            )

        return anchor_xml

    new_xml_str = anchor_pattern.sub(process_anchor, xml_str)

    return new_xml_str.encode("utf-8")

def adjust_row_heights_xml(

    worksheet_xml: bytes,

    sheet_values: dict[str, object],

    workbook,

    sheet_name: str,

) -> bytes:

    """Điều chỉnh chiều cao hàng trong bảng dữ liệu cho đồng bộ theo nội dung.

    - Hàng có dữ liệu: tự động tính chiều cao theo số dòng text (wrap text).

    - Hàng trống trong vùng bảng: chiều cao thấp cố định (~14.4pt).

    - Giữ nguyên chiều cao các hàng ngoài vùng bảng (header, footer, chữ ký).

    """

    try:

        ws = workbook[sheet_name] if (workbook and sheet_name in workbook.sheetnames) else None

        def get_val(coord: str) -> object:

            """Lấy giá trị ô: ưu tiên replacements, sau đó template."""

            if coord in sheet_values:

                return sheet_values[coord]

            if ws is not None:

                try:

                    return ws[coord].value

                except Exception:

                    return None

            return None

        def get_col_width_chars(col_idx: int) -> float:

            """Lấy chiều rộng cột tính bằng số ký tự."""

            if ws is None:

                return 10.0

            try:

                letter = get_column_letter(col_idx)

                cd = ws.column_dimensions.get(letter)

                if cd and cd.width:

                    return float(cd.width)

            except Exception:

                pass

            return 10.0

        def count_wrap_lines(text: str, col_chars: float) -> int:

            """Ước lượng số dòng khi wrap text trong ô có chiều rộng col_chars."""

            if not text:

                return 1

            lines = text.split("\n")

            total = 0

            for line in lines:

                char_per_line = max(int(col_chars * 1.15), 8)

                total += max(1, math.ceil(len(line) / char_per_line))

            return max(1, total)

        def calc_row_height(row_idx: int, data_cols: list[int]) -> float:

            """Tính chiều cao tối ưu của hàng dựa trên nội dung các cột chính."""

            max_lines = 1

            for c in data_cols:

                coord = f"{get_column_letter(c)}{row_idx}"

                val = get_val(coord)

                if val is None:

                    continue

                text = str(val).strip()

                if not text:

                    continue

                col_w = get_col_width_chars(c)

                lines = count_wrap_lines(text, col_w)

                if lines > max_lines:

                    max_lines = lines

            # Excel row height in points: 1 line ≈ 13.5pt, each extra line +12pt

            if max_lines == 1:

                return 28.5

            return min(13.5 + (max_lines - 1) * 12.5, 80.0)

        # Identify data table row range:

        # Strategy 1: scan sheet_values for col A/B = integer (STT serial numbers)

        # Strategy 2: find header row containing 'STT' in template ws, data starts at header+1

        data_start_row = None

        data_end_row = None

        import re as _re

        # Strategy 1: look in sheet_values for smallest row where A or B is an integer >= 1

        sv_int_rows: list[int] = []

        for coord, val in sheet_values.items():

            m = _re.fullmatch(r"([AB])(\d+)", coord)

            if m and val is not None:

                s = str(val).strip()

                if s.isdigit() and int(s) >= 1:

                    sv_int_rows.append(int(m.group(2)))

        if sv_int_rows:

            data_start_row = min(sv_int_rows)

        # Strategy 2: find header row in template with 'STT' text

        if data_start_row is None and ws is not None:

            for r in range(1, 60):

                for c in range(1, 4):

                    try:

                        v = ws.cell(r, c).value

                        if v and "STT" in str(v).upper():

                            data_start_row = r + 1

                            break

                    except Exception:

                        pass

                if data_start_row:

                    break

        if data_start_row is None:

            return worksheet_xml

        # Find data_end_row: largest row number that has any replacement in sheet_values

        # and is within the data table range (>= data_start_row, < data_start_row+30)

        for coord in sheet_values:

            m_all = _re.fullmatch(r"[A-Z]{1,3}(\d+)", coord)

            if m_all:

                r_num = int(m_all.group(1))

                if r_num >= data_start_row:

                    if data_end_row is None or r_num > data_end_row:

                        data_end_row = r_num

        # Cap at data_start_row + 25 to avoid including footer rows

        if data_end_row is not None and data_end_row > data_start_row + 25:

            data_end_row = data_start_row + 25

        # Fallback: scan XML rows for cells between data_start_row and data_start_row+25

        if data_end_row is None:

            for r in range(data_start_row, data_start_row + 26):

                row_xml_m = re.search(

                    rb'<row\b[^>]*\br="' + str(r).encode() + rb'"[^>]*>.*?</row>',

                    worksheet_xml, re.DOTALL

                )

                if row_xml_m:

                    # Row has cells? Check for <c r="..."> in the row content

                    if re.search(rb'<c\b', row_xml_m.group(0)):

                        data_end_row = r

                else:

                    break

        if data_end_row is None:

            data_end_row = data_start_row

        # Narrow data_end_row: the actual last row with meaningful data

        # (rows replaced with empty string are "empty" rows, stop before them)

        for r in range(data_start_row, data_end_row + 1):

            # Check if any cell in this row in sheet_values has non-empty content

            row_has_content = any(

                sheet_values.get(f"{get_column_letter(c)}{r}", "") not in (None, "", 0)

                for c in range(1, 5)

            )

            if not row_has_content:

                # Check if the row has ANY replacement at all

                row_has_any = any(

                    _re.fullmatch(rf"[A-Z]{{1,3}}{r}", coord)

                    for coord in sheet_values

                )

                if not row_has_any:

                    # No replacements → this row is below the data range

                    data_end_row = r - 1

                    break

        if data_end_row < data_start_row:

            data_end_row = data_start_row

        # Also cover empty rows AFTER data_end_row that are still part of the table

        # (template rows with same large ht as data rows, right after the data section)

        data_row_ht_set: set[str] = set()

        for r in range(data_start_row, data_end_row + 1):

            ht_m = re.search(

                rb'<row\b[^>]*\br="' + str(r).encode() + rb'"[^>]*>',

                worksheet_xml,

            )

            if ht_m:

                ht_v = re.search(rb'\bht="([^"]+)"', ht_m.group(0))

                if ht_v:

                    data_row_ht_set.add(ht_v.group(1).decode())

        empty_after_rows: dict[int, float] = {}

        for r in range(data_end_row + 1, data_end_row + 21):

            ht_m = re.search(

                rb'<row\b[^>]*\br="' + str(r).encode() + rb'"[^>]*>',

                worksheet_xml,

            )

            if not ht_m:

                break

            ht_v = re.search(rb'\bht="([^"]+)"', ht_m.group(0))

            if not ht_v:

                break

            row_ht_str = ht_v.group(1).decode()

            try:

                row_ht_float = float(row_ht_str)

            except ValueError:

                break

            if row_ht_str in data_row_ht_set or row_ht_float >= 25.0:

                empty_after_rows[r] = 14.4  # empty table row → thin height

            else:

                break  # different height = footer row, stop

        # Columns that typically hold wrap content in RP tables

        data_cols = [5, 6, 7, 8, 9, 10, 11, 12]

        # Build row height map: row_idx -> target height (points)

        row_heights_final: dict[int, float] = {}

        for r in range(data_start_row, data_end_row + 1):

            has_data = any(

                str(sheet_values.get(f"{get_column_letter(c)}{r}") or

                 (ws[f"{get_column_letter(c)}{r}"].value if ws else None) or "").strip() != ""

                for c in range(1, 5)

            )

            if has_data:

                row_heights_final[r] = calc_row_height(r, data_cols)

            else:

                row_heights_final[r] = 14.4

        # Merge empty-after-data rows

        for r, ht in empty_after_rows.items():

            row_heights_final[r] = ht

        row_heights = row_heights_final

        # Patch <row ...> tags in worksheet XML

        def patch_row(m: re.Match[bytes]) -> bytes:

            row_tag = m.group(0)

            # Extract r= attribute (row index)

            r_match = re.search(rb'\br="(\d+)"', row_tag)

            if not r_match:

                return row_tag

            r_idx = int(r_match.group(1))

            if r_idx not in row_heights:

                return row_tag

            target_ht = row_heights[r_idx]

            # Remove existing ht and customHeight attributes then inject new ones

            row_tag = re.sub(rb'\s+ht="[^"]*"', b"", row_tag)

            row_tag = re.sub(rb'\s+customHeight="[^"]*"', b"", row_tag)

            ht_bytes = f' ht="{target_ht:.2f}" customHeight="1"'.encode()

            # Insert before closing > of opening tag

            row_tag = re.sub(rb"(/?>\s*)$", ht_bytes + rb"\1", row_tag, count=1)

            return row_tag

        # Match only the opening <row ...> tag (not the content)

        patched = re.sub(rb"<row\b[^>]*>", patch_row, worksheet_xml)

        # Đảm bảo các ô trong bảng dữ liệu (đặc biệt cột B, C) dùng định dạng căn giữa (s=12)

        for r in range(data_start_row, (data_end_row or data_start_row) + 1):

            for col_l in ["B", "C"]:

                coord_bytes = f"{col_l}{r}".encode("ascii")

                patched = re.sub(

                    rb'(<c\b[^>]*\br="' + coord_bytes + rb'"[^>]*\bs=")13(")',

                    rb'\g<1>12\g<2>',

                    patched,

                )

                patched = re.sub(

                    rb'(<c\b[^>]*\bs=")13("[^>]*\br="' + coord_bytes + rb'")',

                    rb'\g<1>12\g<2>',

                    patched,

                )

        return patched

    except Exception:

        return worksheet_xml

def apply_station_code_to_template_file(file_path: Path, station_code: str) -> bool:

    """Tự nhận diện và cập nhật tên trạm trong file biểu mẫu theo mã vị trí."""

    try:

        workbook = load_workbook(

            file_path,

            read_only=False,

            data_only=False,

            keep_vba=file_path.suffix.lower() == ".xlsm",

            keep_links=True,

        )

    except Exception:

        return False

    station_updates: dict[tuple[str, str], str] = {}

    for sheet in workbook.worksheets:

        for row in sheet.iter_rows():

            for cell in row:

                if cell.data_type != "f" and is_station_cell(cell.value):

                    updated_val = update_station_text(str(cell.value), station_code)

                    if updated_val != str(cell.value):

                        station_updates[(sheet.title, cell.coordinate)] = updated_val

    if not station_updates:

        workbook.close()

        return False

    with zipfile.ZipFile(file_path, "r") as archive:

        paths = worksheet_paths(archive)

    workbook.close()

    temp_path = file_path.with_name(file_path.name + ".tmp_station")

    try:

        with zipfile.ZipFile(file_path, "r") as original:

            sheet_updates: dict[str, bytes] = {}

            for (sheet_name, coordinate), text_val in station_updates.items():

                part = paths.get(sheet_name)

                if not part:

                    continue

                if part not in sheet_updates:

                    sheet_updates[part] = original.read(part)

                sheet_updates[part] = replace_cell_content(

                    sheet_updates[part], coordinate, text_val

                )

            with zipfile.ZipFile(

                temp_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6

            ) as output:

                for item in original.infolist():

                    data = sheet_updates.get(item.filename)

                    if data is None:

                        data = original.read(item.filename)

                    output.writestr(item, data)

        temp_path.replace(file_path)

        return True

    finally:

        if temp_path.exists():

            temp_path.unlink()

def fix_all_templates_for_site(location: str, custom_code: str = "") -> dict[str, object]:

    """Tự động quét và sửa đổi ký tự địa điểm trạm trên toàn bộ biểu mẫu RFI và RP của một vị trí."""

    location = validate_site_code(location, require_existing=True)

    station_code = custom_code.strip() if custom_code.strip() else station_code_for_location(location)

    results: list[dict[str, object]] = []

    total_updated = 0

    for family in ("rfi", "rp"):

        folder = site_template_dir(location, family)

        if not folder.is_dir():

            continue

        for candidate in folder.iterdir():

            if candidate.suffix.lower() not in {".xlsx", ".xlsm"}:

                continue

            if candidate.name.startswith("FORM CHOOSE"):

                continue

            updated = apply_station_code_to_template_file(candidate, station_code)

            if updated:

                total_updated += 1

            results.append({

                "family": family.upper(),

                "file": candidate.name,

                "updated": updated,

            })

    return {

        "location": location,

        "station_code": station_code,

        "total_files": len(results),

        "updated_files": total_updated,

        "details": results,

    }

def apply_drawings_to_rp_template_file(

    file_path: Path, drawing_no: str, weld_map_no: str

) -> bool:

    """Tự động nhận diện và cập nhật số bản vẽ và số bản vẽ mối hàn vào file biểu mẫu RP."""

    try:

        workbook = load_workbook(

            file_path,

            read_only=False,

            data_only=False,

            keep_vba=file_path.suffix.lower() == ".xlsm",

            keep_links=True,

        )

    except Exception:

        return False

    drawing_updates: dict[tuple[str, str], str] = {}

    for sheet in workbook.worksheets:

        for row in sheet.iter_rows():

            for cell in row:

                if cell.data_type != "f" and cell.value:

                    new_val, changed = update_drawing_cell_value(cell.value, drawing_no, weld_map_no)

                    if changed:

                        drawing_updates[(sheet.title, cell.coordinate)] = new_val

    if not drawing_updates:

        workbook.close()

        return False

    with zipfile.ZipFile(file_path, "r") as archive:

        paths = worksheet_paths(archive)

    workbook.close()

    temp_path = file_path.with_name(file_path.name + ".tmp_drawing")

    try:

        with zipfile.ZipFile(file_path, "r") as original:

            sheet_updates: dict[str, bytes] = {}

            for (sheet_name, coordinate), text_val in drawing_updates.items():

                part = paths.get(sheet_name)

                if not part:

                    continue

                if part not in sheet_updates:

                    sheet_updates[part] = original.read(part)

                sheet_updates[part] = replace_cell_content(

                    sheet_updates[part], coordinate, text_val

                )

            with zipfile.ZipFile(

                temp_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6

            ) as output:

                for item in original.infolist():

                    data = sheet_updates.get(item.filename)

                    if data is None:

                        data = original.read(item.filename)

                    output.writestr(item, data)

        temp_path.replace(file_path)

        return True

    finally:

        if temp_path.exists():

            temp_path.unlink()

def update_all_rp_drawings_for_site(

    location: str,

    drawing_no: str,

    weld_map_no: str,

    engineer_name: str | None = None,

) -> dict[str, object]:

    """Cập nhật số bản vẽ cho tất cả biểu mẫu RP của một vị trí và lưu cấu hình."""

    location = validate_site_code(location, require_existing=True)

    saved_cfg = save_site_config(location, drawing_no, weld_map_no, engineer_name)

    folder = site_template_dir(location, "rp")

    results = []

    total_updated = 0

    if folder.is_dir():

        for candidate in sorted(folder.iterdir()):

            if candidate.suffix.lower() not in {".xlsx", ".xlsm"}:

                continue

            if candidate.name.startswith("FORM CHOOSE"):

                continue

            updated = apply_drawings_to_rp_template_file(

                candidate, saved_cfg["drawing_no"], saved_cfg["weld_map_no"]

            )

            if updated:

                total_updated += 1

            results.append({"file": candidate.name, "updated": updated})

    return {

        "location": location,

        "drawing_no": saved_cfg["drawing_no"],

        "weld_map_no": saved_cfg["weld_map_no"],

        "engineer_name": saved_cfg.get("engineer_name", default_engineer_for_location(location)),

        "total_files": len(results),

        "updated_files": total_updated,

        "details": results,

    }

RP_ASSIGNMENT_COLUMNS = {

    "FUR": "AJ",

    "VSR": "AR",

    "NDT": "BG",

    "MTR": "EB",

    "PPP": "EI",

    "SPI": "CM",

    "CAI": "CR",

    "FTI": "CW",

    "HTI": "DB",

}

RP_SELECTORS = {

    "ICR": "J7",

    "MTR": "J7",

    "PPP": "J6",

    "FUR": "O7",

    "VSR": "M7",

    "DIC": "O8",

    "SPI": "J6",

    "CAI": "J6",

    "FTI": "J6",

    "HTI": "J7",

}

CELL_REFERENCE_RE = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$", re.IGNORECASE)

RANGE_REFERENCE_RE = re.compile(

    r"^\$?([A-Z]{1,3})\$?(\d+)?(?::\$?([A-Z]{1,3})\$?(\d+)?)?$",

    re.IGNORECASE,

)

FORMULA_TOKEN_RE = re.compile(

    r'\s*(?:(?P<string>"(?:[^"]|"")*")'

    r"|(?P<quoted>'(?:[^']|'')*'!)"

    r"|(?P<sheet>(?:\[\d+\])?[A-Za-z_][A-Za-z0-9_. ]*!)"

    r"|(?P<reference>\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?"

    r"|\$?[A-Z]{1,3}:\$?[A-Z]{1,3})"

    r"|(?P<number>(?:\d+(?:\.\d*)?|\.\d+)%?)"

    r"|(?P<operator><>|<=|>=|[(),+\-*/^&=<>])"

    r"|(?P<identifier>[A-Za-z_][A-Za-z0-9_.]*))",

    re.IGNORECASE,

)

class FormulaError(Exception):

    pass

@dataclass(frozen=True)

class FormulaReference:

    sheet: str | None

    start_column: int

    start_row: int

    end_column: int

    end_row: int

    is_range: bool

    is_external: bool

@dataclass

class FormulaRange:

    values: list[list[object]]

    row_count: int

    column_count: int

    def flattened(self) -> list[object]:

        return [value for row in self.values for value in row]

def excel_truthy(value: object) -> bool:

    if isinstance(value, FormulaRange):

        value = value.flattened()[0] if value.flattened() else None

    if value is None or value == "":

        return False

    if isinstance(value, str):

        return value.casefold() not in {"false", "0"}

    return bool(value)

def excel_number(value: object) -> float:

    if value is None or value == "":

        return 0.0

    if isinstance(value, bool):

        return 1.0 if value else 0.0

    if isinstance(value, (int, float)):

        return float(value)

    try:

        return float(str(value).replace(",", ""))

    except ValueError as exc:

        raise FormulaError(f"Không thể chuyển {value!r} thành số.") from exc

def normalized_sheet_name(name: str | None) -> str:

    if not name:

        return ""

    value = name.strip().strip("'").replace("''", "'")

    value = re.sub(r"^\[\d+\]", "", value)

    return " ".join(value.split()).casefold()

def formula_text(value: object) -> str:

    if isinstance(value, str):

        return value

    text = getattr(value, "text", None)

    return text if isinstance(text, str) else str(value)

def iter_formula_cells(sheet):

    return (cell for cell in sheet._cells.values() if cell.data_type == "f")

class RPFormulaEngine:

    def __init__(

        self,

        formula_workbook,

        form_rows: dict[str, dict[int, object]],

        external_tables: dict[str, list[tuple[int, tuple[object, ...]]]],

        helper_values: dict[tuple[str, str], object],

        selected_id: str,

    ):

        self.workbook = formula_workbook

        self.form_rows = form_rows

        self.external_tables = {

            normalized_sheet_name(name): rows for name, rows in external_tables.items()

        }

        self.external_table_map = {

            s: {r: vals for r, vals in rows} for s, rows in self.external_tables.items()

        }

        self.helper_values = helper_values

        self.selected_id = selected_id

        self.cell_overrides: dict[tuple[str, str], object] = dict(helper_values)

        self.evaluated: dict[tuple[str, str], object] = {}

        self.active: set[tuple[str, str]] = set()

        self.range_cache: dict[tuple[object, ...], FormulaRange] = {}

        self.sheet_dimensions = {

            sheet.title: (

                max((row for row, _ in sheet._cells), default=0),

                max((column for _, column in sheet._cells), default=0),

            )

            for sheet in formula_workbook.worksheets

        }

        self.current_cell = ""

    def evaluate_cell(self, sheet_name: str, coordinate: str) -> object:

        key = (sheet_name, coordinate)

        if key in self.cell_overrides:

            return self.cell_overrides[key]

        if key in self.evaluated:

            return self.evaluated[key]

        if key in self.active:

            raise FormulaError(f"Phát hiện tham chiếu vòng tại {sheet_name}!{coordinate}.")

        if sheet_name not in self.workbook.sheetnames:

            raise FormulaError(f"Không tìm thấy sheet {sheet_name}.")

        cell = self.workbook[sheet_name][coordinate]

        if cell.data_type != "f":

            return cell.value

        self.active.add(key)

        previous_cell = self.current_cell

        self.current_cell = f"{sheet_name}!{coordinate}"

        try:

            value = self.evaluate_formula(formula_text(cell.value), sheet_name)

            self.evaluated[key] = value

            return value

        except FormulaError as exc:

            raise FormulaError(f"{sheet_name}!{coordinate}: {exc}") from exc

        finally:

            self.current_cell = previous_cell

            self.active.discard(key)

    def _tokenize(self, formula: str) -> list[tuple[str, str]]:

        expression = formula[1:] if formula.startswith("=") else formula

        tokens: list[tuple[str, str]] = []

        position = 0

        while position < len(expression):

            match = FORMULA_TOKEN_RE.match(expression, position)

            if match is None:

                if expression[position:].strip() == "":

                    break

                raise FormulaError(f"Công thức không được hỗ trợ: {formula}")

            kind = match.lastgroup

            if kind:

                tokens.append((kind, match.group(kind)))

            position = match.end()

        tokens.append(("eof", ""))

        return tokens

    def evaluate_formula(self, formula: str, current_sheet: str) -> object:

        parser = _FormulaParser(self, self._tokenize(formula), current_sheet)

        expression = parser.parse_expression()

        if parser.peek()[0] != "eof":

            raise FormulaError(f"Công thức còn phần chưa xử lý: {formula}")

        return self._evaluate(expression, current_sheet)

    def _parse_reference(

        self, reference: str, sheet: str | None, external: bool

    ) -> FormulaReference:

        match = RANGE_REFERENCE_RE.fullmatch(reference)

        if match is None:

            raise FormulaError(f"Tham chiếu ô không hợp lệ: {reference}")

        start_column = column_index_from_string(match.group(1).upper())

        end_column = column_index_from_string((match.group(3) or match.group(1)).upper())

        start_row = int(match.group(2) or "1")

        end_row = int(match.group(4) or (str(start_row) if match.group(2) else "1048576"))

        return FormulaReference(

            sheet=sheet,

            start_column=start_column,

            start_row=start_row,

            end_column=end_column,

            end_row=end_row,

            is_range=":" in reference,

            is_external=external,

        )

    def _get_cell_value(

        self, reference: FormulaReference, row: int, column: int, current_sheet: str

    ) -> object:

        sheet_name = reference.sheet or current_sheet

        if reference.is_external:

            t_map = self.external_table_map.get(normalized_sheet_name(sheet_name))

            if t_map is None:

                raise FormulaError(f"Không nhận diện được file nguồn cho sheet {sheet_name}.")

            values = t_map.get(row)

            if values is None or column < 1 or column > len(values):

                return None

            return values[column - 1]

        if sheet_name not in self.workbook.sheetnames:

            raise FormulaError(f"Không tìm thấy sheet {sheet_name}.")

        max_row, max_column = self.sheet_dimensions[sheet_name]

        if row < 1 or row > max_row or column < 1 or column > max_column:

            return None

        coordinate = f"{get_column_letter(column)}{row}"

        return self.evaluate_cell(sheet_name, coordinate)

    def _resolve_reference(

        self, reference: FormulaReference, current_sheet: str

    ) -> object:

        if not reference.is_range:

            return self._get_cell_value(

                reference, reference.start_row, reference.start_column, current_sheet

            )

        sheet_name = reference.sheet or current_sheet

        cache_key = (

            sheet_name.casefold(),

            reference.start_column,

            reference.start_row,

            reference.end_column,

            reference.end_row,

            reference.is_external,

        )

        cached = self.range_cache.get(cache_key)

        if cached is not None:

            return cached

        if reference.is_external:

            table = self.external_tables.get(normalized_sheet_name(sheet_name))

            if table is None:

                raise FormulaError(f"Không nhận diện được file nguồn cho sheet {sheet_name}.")

            last_row = min(reference.end_row, table[-1][0] if table else 0)

            t_map = self.external_table_map.get(normalized_sheet_name(sheet_name), {})

            rows: list[list[object]] = []

            if last_row >= reference.start_row:

                col_range = range(reference.start_column, reference.end_column + 1)

                for row_number in range(reference.start_row, last_row + 1):

                    vals = t_map.get(row_number)

                    if vals is None:

                        rows.append([None] * len(col_range))

                    else:

                        rows.append([

                            vals[c - 1] if 1 <= c <= len(vals) else None

                            for c in col_range

                        ])

        else:

            if sheet_name not in self.workbook.sheetnames:

                raise FormulaError(f"Không tìm thấy sheet {sheet_name}.")

            last_row = min(reference.end_row, self.sheet_dimensions[sheet_name][0])

            rows: list[list[object]] = []

            if last_row >= reference.start_row:

                for row_number in range(reference.start_row, last_row + 1):

                    rows.append(

                        [

                            self._get_cell_value(reference, row_number, column, current_sheet)

                            for column in range(reference.start_column, reference.end_column + 1)

                        ]

                    )

        value = FormulaRange(

            values=rows,

            row_count=max(0, reference.end_row - reference.start_row + 1),

            column_count=max(1, reference.end_column - reference.start_column + 1),

        )

        self.range_cache[cache_key] = value

        return value

    def _evaluate(self, node: object, current_sheet: str) -> object:

        kind = node[0]

        if kind == "literal":

            return node[1]

        if kind == "reference":

            return self._resolve_reference(node[1], current_sheet)

        if kind == "unary":

            value = excel_number(self._evaluate(node[2], current_sheet))

            return value if node[1] == "+" else -value

        if kind == "binary":

            left = self._evaluate(node[2], current_sheet)

            right = self._evaluate(node[3], current_sheet)

            return self._evaluate_binary(node[1], left, right)

        if kind == "function":

            return self._evaluate_function(node[1], node[2], current_sheet)

        raise FormulaError("Không thể diễn giải công thức.")

    def _evaluate_binary(self, operator: str, left: object, right: object) -> object:

        if operator in {"=", "<>", "<", ">", "<=", ">="} and (

            isinstance(left, FormulaRange) or isinstance(right, FormulaRange)

        ):

            left_values = left.flattened() if isinstance(left, FormulaRange) else None

            right_values = right.flattened() if isinstance(right, FormulaRange) else None

            count = len(left_values if left_values is not None else right_values or [])

            results = []

            for index in range(count):

                left_value = left_values[index] if left_values is not None else left

                right_value = right_values[index] if right_values is not None else right

                results.append(self._evaluate_binary(operator, left_value, right_value))

            return FormulaRange(

                values=[[value] for value in results],

                row_count=len(results),

                column_count=1,

            )

        if isinstance(left, FormulaRange):

            left = left.flattened()[0] if left.flattened() else None

        if isinstance(right, FormulaRange):

            right = right.flattened()[0] if right.flattened() else None

        if operator == "&":

            return cell_text(left) + cell_text(right)

        if operator in {"=", "<>", "<", ">", "<=", ">="}:

            if left is None:

                left = ""

            if right is None:

                right = ""

            if operator in {"=", "<>"}:

                equal = str(left).casefold() == str(right).casefold() if isinstance(left, str) or isinstance(right, str) else left == right

                return equal if operator == "=" else not equal

            if isinstance(left, str) or isinstance(right, str):

                left, right = str(left or ""), str(right or "")

            else:

                left, right = excel_number(left), excel_number(right)

            return {

                "<": left < right,

                ">": left > right,

                "<=": left <= right,

                ">=": left >= right,

            }[operator]

        a, b = excel_number(left), excel_number(right)

        if operator == "+":

            return a + b

        if operator == "-":

            return a - b

        if operator == "*":

            return a * b

        if operator == "/":

            if b == 0:

                raise FormulaError("Chia cho 0.")

            return a / b

        if operator == "^":

            return a**b

        raise FormulaError(f"Toán tử không được hỗ trợ: {operator}")

    @staticmethod

    def _flatten(value: object) -> list[object]:

        if isinstance(value, FormulaRange):

            return value.flattened()

        if isinstance(value, list):

            result: list[object] = []

            for item in value:

                result.extend(RPFormulaEngine._flatten(item))

            return result

        return [value]

    @staticmethod

    def _scalar(value: object) -> object:

        if isinstance(value, FormulaRange):

            values = value.flattened()

            return values[0] if values else None

        if isinstance(value, list):

            return value[0] if value else None

        return value

    def _evaluate_function(

        self, name: str, arguments: list[object], current_sheet: str

    ) -> object:

        function = name.upper()

        if function.startswith("_XLFN."):

            function = function.rsplit(".", 1)[-1]

        if function == "IF":

            if len(arguments) not in {2, 3}:

                raise FormulaError("IF cần 2 hoặc 3 tham số.")

            condition = excel_truthy(self._evaluate(arguments[0], current_sheet))

            return self._evaluate(arguments[1 if condition else 2], current_sheet) if condition or len(arguments) == 3 else False

        if function == "IFERROR":

            try:

                return self._evaluate(arguments[0], current_sheet)

            except (FormulaError, ZeroDivisionError, ValueError, TypeError):

                return self._evaluate(arguments[1], current_sheet) if len(arguments) > 1 else ""

        if function == "CHOOSE":

            index = int(excel_number(self._evaluate(arguments[0], current_sheet)))

            if index < 1 or index >= len(arguments):

                raise FormulaError("CHOOSE có chỉ số ngoài phạm vi.")

            return self._evaluate(arguments[index], current_sheet)

        if function in {"ROWS", "COLUMNS"}:

            if not arguments:

                raise FormulaError(f"{function} cần một vùng.")

            arg = arguments[0]

            if isinstance(arg, tuple) and arg[0] == "reference":

                ref = arg[1]

                return max(1, ref.end_row - ref.start_row + 1) if function == "ROWS" else max(1, ref.end_column - ref.start_column + 1)

            val = self._evaluate(arg, current_sheet)

            if isinstance(val, FormulaRange):

                return val.row_count if function == "ROWS" else val.column_count

            return 1

        values = [self._evaluate(argument, current_sheet) for argument in arguments]

        if function in {"XLOOKUP", "VLOOKUP"}:

            if function == "XLOOKUP":

                if len(values) < 3:

                    raise FormulaError("XLOOKUP cần ít nhất 3 tham số.")

                lookup_value = self._scalar(values[0])

                lookup_values = self._flatten(values[1])

                return_values = self._flatten(values[2])

                if lookup_value not in (None, "") and str(lookup_value).strip():

                    for index, candidate in enumerate(lookup_values):

                        if str(candidate).strip().casefold() == str(lookup_value).strip().casefold():

                            return return_values[index] if index < len(return_values) else None

                if len(values) > 3:

                    return values[3]

                raise FormulaError(f"Không tìm thấy {lookup_value!r} trong XLOOKUP.")

            if len(values) < 3:

                raise FormulaError("VLOOKUP cần ít nhất 3 tham số.")

            lookup_value = self._scalar(values[0])

            if lookup_value in (None, "") or not str(lookup_value).strip():

                raise FormulaError("Không tìm thấy ô rỗng trong VLOOKUP.")

            table = values[1]

            if not isinstance(table, FormulaRange):

                raise FormulaError("Đối số bảng của VLOOKUP phải là một vùng.")

            return_column = int(excel_number(values[2])) - 1

            for row in table.values:

                if row and str(row[0]).strip().casefold() == str(lookup_value).strip().casefold():

                    return row[return_column] if return_column < len(row) else None

            raise FormulaError(f"Không tìm thấy {lookup_value!r} trong VLOOKUP.")

        if function == "MATCH":

            if len(values) < 2:

                raise FormulaError("MATCH cần ít nhất 2 tham số.")

            lookup_value = self._scalar(values[0])

            if lookup_value in (None, "") or not str(lookup_value).strip():

                raise FormulaError("Không tìm thấy ô rỗng trong MATCH.")

            lookup_values = self._flatten(values[1])

            match_type = int(excel_number(values[2])) if len(values) > 2 else 1

            if match_type != 0:

                raise FormulaError("Hiện chỉ hỗ trợ MATCH chính xác (tham số 0).")

            for index, candidate in enumerate(lookup_values, start=1):

                if str(candidate).strip().casefold() == str(lookup_value).strip().casefold():

                    return index

            raise FormulaError(f"Không tìm thấy {lookup_value!r} trong MATCH.")

        if function == "INDEX":

            if len(values) < 2:

                raise FormulaError("INDEX cần ít nhất 2 tham số.")

            array = values[0]

            row_number = int(excel_number(values[1]))

            column_number = int(excel_number(values[2])) if len(values) > 2 else 1

            if isinstance(array, FormulaRange):

                if row_number < 1 or row_number > len(array.values):

                    raise FormulaError("INDEX không tìm thấy hàng.")

                row = array.values[row_number - 1]

                if column_number == 0:

                    return row

                if column_number > len(row):

                    raise FormulaError("INDEX không tìm thấy cột.")

                return row[column_number - 1]

            flattened = self._flatten(array)

            if row_number < 1 or row_number > len(flattened):

                raise FormulaError("INDEX không tìm thấy hàng.")

            return flattened[row_number - 1]

        if function == "COUNTA":

            return sum(value is not None and value != "" for item in values for value in self._flatten(item))

        if function == "COUNTIF":

            if len(values) < 2:

                raise FormulaError("COUNTIF cần 2 tham số.")

            criterion = str(self._scalar(values[1])).strip().casefold()

            return sum(

                str(value).strip().casefold() == criterion

                for value in self._flatten(values[0])

                if value is not None

            )

        if function == "SUBTOTAL":

            if len(values) < 2 or int(excel_number(values[0])) not in {3, 103}:

                raise FormulaError("Hiện chỉ hỗ trợ SUBTOTAL(3, vùng).")

            return sum(value is not None and value != "" for item in values[1:] for value in self._flatten(item))

        if function == "FILTER":

            if len(values) < 2:

                raise FormulaError("FILTER cần ít nhất 2 tham số.")

            candidates = self._flatten(values[0])

            flags = self._flatten(values[1])

            matches = [item for item, flag in zip(candidates, flags) if excel_truthy(flag)]

            if matches:

                return matches

            return self._flatten(values[2]) if len(values) > 2 else []

        if function == "TRIM":

            value = values[0]

            if isinstance(value, FormulaRange):

                return FormulaRange(

                    values=[

                        [" ".join(str(item or "").split()) for item in row]

                        for row in value.values

                    ],

                    row_count=value.row_count,

                    column_count=value.column_count,

                )

            return " ".join(str(self._scalar(value) or "").split())

        if function == "SUBSTITUTE":

            text = str(self._scalar(values[0]) or "")

            old = str(self._scalar(values[1]) or "")

            new = str(self._scalar(values[2]) or "")

            instance = int(excel_number(values[3])) if len(values) > 3 else 0

            if instance <= 0:

                return text.replace(old, new)

            split = text.split(old)

            return old.join(split[:instance]) + new + old.join(split[instance:]) if len(split) > instance else text

        if function == "RAND":

            return random.random()

        if function == "RANDBETWEEN":

            if len(values) < 2:

                raise FormulaError("RANDBETWEEN cần 2 tham số.")

            low_val = values[0]

            high_val = values[1]

            if low_val in (None, "") or high_val in (None, ""):

                return ""

            low = int(round(excel_number(low_val)))

            high = int(round(excel_number(high_val)))

            if low > high:

                low, high = high, low

            return random.randint(low, high)

        if function == "ROUND":

            if len(values) < 2:

                raise FormulaError("ROUND cần 2 tham số.")

            precision = int(excel_number(values[1]))

            return round(excel_number(values[0]), precision)

        if function == "SUM":

            return sum(

                excel_number(value)

                for item in values

                for value in self._flatten(item)

                if value not in (None, "")

            )

        if function == "ABS":

            return abs(excel_number(values[0]))

        if function == "AND":

            return all(excel_truthy(value) for item in values for value in self._flatten(item))

        if function == "OR":

            return any(excel_truthy(value) for item in values for value in self._flatten(item))

        if function == "LEFT":

            text = str(self._scalar(values[0]) or "")

            count = int(excel_number(values[1])) if len(values) > 1 else 1

            return text[:count]

        if function == "MID":

            if len(values) < 3:

                raise FormulaError("MID cần 3 tham số.")

            text = str(self._scalar(values[0]) or "")

            start = max(1, int(excel_number(values[1]))) - 1

            length = int(excel_number(values[2]))

            return text[start : start + length]

        if function == "LEN":

            return len(str(self._scalar(values[0]) or ""))

        if function == "VALUE":

            if not values:

                raise FormulaError("VALUE cần 1 tham số.")

            return excel_number(values[0])

        if function == "TEXTBEFORE":

            text = str(self._scalar(values[0]) or "")

            delimiter = str(self._scalar(values[1]) or "")

            pos = text.find(delimiter)

            return text[:pos] if pos >= 0 else text

        if function == "TEXTAFTER":

            text = str(self._scalar(values[0]) or "")

            delimiter = str(self._scalar(values[1]) or "")

            pos = text.find(delimiter)

            return text[pos + len(delimiter) :] if pos >= 0 else text

        if function == "SEARCH":

            needle = str(self._scalar(values[0]))

            text = str(self._scalar(values[1]))

            start = int(excel_number(values[2])) - 1 if len(values) > 2 else 0

            found = text.casefold().find(needle.casefold(), start)

            if found < 0:

                raise FormulaError(f"Không tìm thấy {needle!r}.")

            return found + 1

        raise FormulaError(f"Chưa hỗ trợ hàm Excel {name}.")

def find_random_measurement_formulas(workbook) -> set[tuple[str, str]]:

    formulas: dict[tuple[str, str], str] = {}

    for sheet in workbook.worksheets:

        for cell in iter_formula_cells(sheet):

            formulas[(sheet.title, cell.coordinate)] = formula_text(cell.value)

    preserved = {

        key for key, formula in formulas.items() if RANDOM_MEASUREMENT_FORMULA.search(formula)

    }

    changed = True

    while changed:

        changed = False

        for key, formula in formulas.items():

            if key in preserved:

                continue

            expression = re.sub(r'"(?:""|[^"])*"', '""', formula)

            for match in FORMULA_CELL_REFERENCE.finditer(expression):

                reference = match.group(3).replace("$", "").upper()

                sheet_reference = match.group(1) or match.group(2)

                target_sheet = key[0] if not sheet_reference else None

                if sheet_reference:

                    normalized_reference = normalized_sheet_name(sheet_reference)

                    for sheet_name in workbook.sheetnames:

                        if normalized_sheet_name(sheet_name) == normalized_reference:

                            target_sheet = sheet_name

                            break

                if target_sheet is not None and (target_sheet, reference) in preserved:

                    preserved.add(key)

                    changed = True

                    break

    return preserved

class _FormulaParser:

    PRECEDENCE = {

        "=": 5,

        "<>": 5,

        "<": 5,

        ">": 5,

        "<=": 5,

        ">=": 5,

        "&": 10,

        "+": 20,

        "-": 20,

        "*": 30,

        "/": 30,

        "^": 40,

    }

    def __init__(

        self, engine: RPFormulaEngine, tokens: list[tuple[str, str]], current_sheet: str

    ):

        self.engine = engine

        self.tokens = tokens

        self.current_sheet = current_sheet

        self.index = 0

    def peek(self) -> tuple[str, str]:

        return self.tokens[self.index]

    def consume(self) -> tuple[str, str]:

        token = self.peek()

        self.index += 1

        return token

    def parse_expression(self, minimum_precedence: int = 0) -> object:

        token_type, value = self.consume()

        if token_type == "operator" and value in {"+", "-"}:

            left = ("unary", value, self.parse_expression(35))

        elif token_type == "operator" and value == "(":

            left = self.parse_expression()

            if self.consume() != ("operator", ")"):

                raise FormulaError("Thiếu dấu ngoặc đóng.")

        elif token_type == "string":

            left = ("literal", value[1:-1].replace('""', '"'))

        elif token_type == "number":

            number = float(value[:-1]) / 100 if value.endswith("%") else float(value)

            left = ("literal", int(number) if number.is_integer() else number)

        elif token_type == "sheet" or token_type == "quoted":

            prefix = value[:-1]

            sheet_name = prefix.strip("'").replace("''", "'")

            external = bool(re.match(r"^\[\d+\]", sheet_name))

            reference_type, reference = self.consume()

            if reference_type != "reference":

                raise FormulaError("Thiếu ô hoặc vùng sau tên sheet.")

            left = (

                "reference",

                self.engine._parse_reference(reference, sheet_name, external),

            )

        elif token_type == "reference":

            left = (

                "reference",

                self.engine._parse_reference(value, None, False),

            )

        elif token_type == "identifier":

            upper = value.upper()

            if upper in {"TRUE", "FALSE"}:

                left = ("literal", upper == "TRUE")

            elif self.peek() == ("operator", "("):

                self.consume()

                arguments: list[object] = []

                if self.peek() != ("operator", ")"):

                    while True:

                        arguments.append(self.parse_expression())

                        if self.peek() != ("operator", ","):

                            break

                        self.consume()

                if self.consume() != ("operator", ")"):

                    raise FormulaError(f"Thiếu dấu ngoặc sau hàm {value}.")

                left = ("function", value, arguments)

            else:

                raise FormulaError(f"Tên hoặc hàm không được hỗ trợ: {value}")

        else:

            raise FormulaError(f"Token không được hỗ trợ: {value}")

        while self.peek()[0] == "operator":

            operator = self.peek()[1]

            precedence = self.PRECEDENCE.get(operator, -1)

            if precedence < minimum_precedence:

                break

            self.consume()

            next_minimum = precedence if operator == "^" else precedence + 1

            right = self.parse_expression(next_minimum)

            left = ("binary", operator, left, right)

        return left

def load_rp_source_tables(

    location: str = DEFAULT_SITE,

) -> dict[str, list[tuple[int, tuple[object, ...]]]]:

    location = validate_site_code(location, require_existing=True)

    tables: dict[str, list[tuple[int, tuple[object, ...]]]] = {}

    loaded_sources: dict[str, object] = {}

    controlog_path = source_path("controlog", location)

    if not controlog_path.is_file():

        raise UserError(

            f"Vị trí {location} chưa có file nguồn {source_filename('controlog', location)}."

        )

    for key, sheet_prefix in (("controlog", "control log"), ("controlog", "pipe ")):

        workbook = loaded_sources.get(key)

        if workbook is None:

            workbook = load_workbook(

                source_path(key, location), read_only=True, data_only=True

            )

            loaded_sources[key] = workbook

        sheet_name = next(

            (

                name

                for name in workbook.sheetnames

                if normalized_sheet_name(name).startswith(sheet_prefix.strip())

            ),

            None,

        )

        if sheet_name is None:

            for opened in loaded_sources.values():

                opened.close()

            raise UserError(

                f"File Controlog của {location} thiếu sheet bắt đầu bằng "

                f"'{sheet_prefix.strip()}'."

            )

        sheet = workbook[sheet_name]

        rows = [

            (row_number, tuple(values))

            for row_number, values in enumerate(

                sheet.iter_rows(values_only=True), start=1

            )

        ]

        tables[sheet_name] = rows

        # Register aliases so formula engine can resolve sheet references.

        # Use location-specific aliases AND the legacy TTC15A aliases (which

        # the RP templates embed as external-link sheet names).

        if sheet_prefix == "control log":

            tables[f"Control Log  {location}"] = rows

            tables[f"Control Log {location}"] = rows

            # Legacy aliases for templates that still reference TTC15A

            if location != DEFAULT_SITE:

                tables["Control Log  TTC15A"] = rows

                tables["Control Log TTC15A"] = rows

        else:

            tables[f"Pipe {location}"] = rows

            if location != DEFAULT_SITE:

                tables["Pipe TTC15A"] = rows

    form_path = source_path("rfi_data", location)

    if not form_path.is_file():

        for opened in loaded_sources.values():

            opened.close()

        raise UserError(

            f"Vị trí {location} chưa có file nguồn {source_filename('rfi_data', location)}."

        )

    form_book = load_workbook(form_path, read_only=True, data_only=True)

    try:

        form_name = form_sheet_name(form_book)

        if form_name is None:

            form_book.close()

            for opened in loaded_sources.values():

                opened.close()

            raise UserError(

                f"Không tìm thấy sheet FORM hoặc tiêu đề RFI Number trong file của {location}."

            )

        form_sheet = form_book[form_name]

        tables["FORM"] = [

            (row_number, tuple(values))

            for row_number, values in enumerate(form_sheet.iter_rows(values_only=True), start=1)

        ]

    finally:

        form_book.close()

        for opened in loaded_sources.values():

            opened.close()

    return tables

_RP_SOURCE_TABLES_CACHE: dict[str, tuple[float, float, dict[str, list[tuple[int, tuple[object, ...]]]]]] = {}

def get_cached_rp_source_tables(

    location: str = DEFAULT_SITE,

) -> dict[str, list[tuple[int, tuple[object, ...]]]]:

    location = validate_site_code(location, require_existing=True)

    c_path = source_path("controlog", location)

    r_path = source_path("rfi_data", location)

    if not c_path.is_file() or not r_path.is_file():

        return load_rp_source_tables(location)

    c_mtime = c_path.stat().st_mtime

    r_mtime = r_path.stat().st_mtime

    cached = _RP_SOURCE_TABLES_CACHE.get(location)

    if cached and cached[0] == c_mtime and cached[1] == r_mtime:

        return cached[2]

    tables = load_rp_source_tables(location)

    _RP_SOURCE_TABLES_CACHE[location] = (c_mtime, r_mtime, tables)

    return tables

def build_rp_helper_values(

    workbook,

    report_type: str,

    control_table: list[tuple[int, tuple[object, ...]]],

    pipe_table: list[tuple[int, tuple[object, ...]]],

) -> dict[tuple[str, str], object]:

    from openpyxl.utils import get_column_letter

    helper_values: dict[tuple[str, str], object] = {}

    for sheet in workbook.worksheets:

        if not sheet.title.casefold().startswith("controlog"):

            continue

        for cell in iter_formula_cells(sheet):

            helper_values[(sheet.title, cell.coordinate)] = None

        grouped: dict[str, int] = {}

        entries: list[tuple[object, object, object]] = []

        if report_type == "ICR":

            table = pipe_table

            key_column = column_index_from_string("P")

            value_column = column_index_from_string("L")

            for row_number, values in table:

                if row_number < 2:

                    continue

                rfi_number = values[key_column - 1] if len(values) >= key_column else None

                pipe_tag = values[value_column - 1] if len(values) >= value_column else None

                if not rfi_number or not pipe_tag:

                    continue

                key = str(rfi_number).strip()

                grouped[key] = grouped.get(key, 0) + 1

                entries.append((key, pipe_tag, f"{key}|{grouped[key]}"))

        else:

            source_column = column_index_from_string(RP_ASSIGNMENT_COLUMNS[report_type])

            joint_column = column_index_from_string("I")

            for row_number, values in control_table:

                if row_number < 4:

                    continue

                rfi_number = values[source_column - 1] if len(values) >= source_column else None

                joint = values[joint_column - 1] if len(values) >= joint_column else None

                if not rfi_number or not joint:

                    continue

                key = str(rfi_number).strip()

                grouped[key] = grouped.get(key, 0) + 1

                entries.append((joint, key, f"{key}|{grouped[key]}"))

        if len(entries) > max(0, sheet.max_row - 1):

            raise UserError(

                f"Sheet phụ {sheet.title} không đủ số dòng cho dữ liệu mới "

                f"({len(entries)} mục)."

            )

        for index, entry in enumerate(entries, start=2):

            for column, value in enumerate(entry, start=1):

                helper_values[(sheet.title, f"{get_column_letter(column)}{index}")] = value

                sheet.cell(row=index, column=column).value = value

        for r in range(len(entries) + 2, sheet.max_row + 1):

            for c in range(1, 4):

                sheet.cell(row=r, column=c).value = None

    return helper_values

def set_export_cell(

    replacements: dict[str, dict[str, object]],

    sheet_name: str,

    coordinate: str,

    value: object,

) -> None:

    replacements.setdefault(sheet_name, {})[coordinate] = value

def format_sheet_numbers(sheets: list[str]) -> str:

    """Định dạng danh sách số sheet thành chuỗi 'A,B/94' hoặc 'A/94' theo yêu cầu.

    Ví dụ: ['39/94', '40/94'] -> '39,40/94'

           ['39/94'] -> '39/94'

    """

    if not sheets:

        return ""

    numerators: list[str] = []

    denominators: list[str] = []

    for s in sheets:

        s = str(s).strip()

        if not s:

            continue

        if "/" in s:

            num_part, denom_part = s.split("/", 1)

            denom = denom_part.strip()

            if denom and denom not in denominators:

                denominators.append(denom)

            for part in num_part.split(","):

                part = part.strip()

                if part and part not in numerators:

                    numerators.append(part)

        else:

            for part in s.split(","):

                part = part.strip()

                if part and part not in numerators:

                    numerators.append(part)

    if not numerators:

        return ""

    def num_key(val: str):

        try:

            return (0, int(val))

        except ValueError:

            return (1, val)

    sorted_nums = sorted(numerators, key=num_key)

    nums_str = ",".join(sorted_nums)

    if denominators:

        return f"{nums_str}/{denominators[0]}"

    return nums_str

def resolve_drawings_for_rfi(

    rfi_id: str,

    report_type: str,

    external_tables: dict[str, list[tuple[int, tuple[object, ...]]]],

    location: str = DEFAULT_SITE,

    fallback_drawing: str = DEFAULT_DRAWING_NO,

    fallback_weld_map: str = DEFAULT_WELD_MAP_NO,

) -> tuple[str, str]:

    """Tự động tra cứu số bản vẽ mối hàn (Cột C, E, D) và số bản vẽ (Cột F, H, G) từ Control Log cho số RFI tương ứng."""

    control_table = None

    for key, rows in external_tables.items():

        if key.lower().startswith("control log"):

            control_table = rows

            if location.lower() in key.lower():

                break

    if not control_table:

        return fallback_drawing, fallback_weld_map

    pipe_table = None

    if report_type == "ICR":

        for key, rows in external_tables.items():

            if key.lower().startswith("pipe"):

                pipe_table = rows

                if location.lower() in key.lower():

                    break

    col_idx = None

    col_letter = RP_ASSIGNMENT_COLUMNS.get(report_type)

    if col_letter:

        col_idx = column_index_from_string(col_letter)

    matched_w_docs: list[str] = []

    matched_w_revs: list[str] = []

    matched_w_sheets: list[str] = []

    matched_d_docs: list[str] = []

    matched_d_revs: list[str] = []

    matched_d_sheets: list[str] = []

    def add_row_match(cvals: tuple[object, ...]):

        c = str(cvals[2] if len(cvals) >= 3 else "" or "").strip()

        d = str(cvals[3] if len(cvals) >= 4 else "" or "").strip()

        e_raw = cvals[4] if len(cvals) >= 5 else None

        e = str(e_raw).strip() if e_raw is not None and str(e_raw).strip() != "" else ""

        f = str(cvals[5] if len(cvals) >= 6 else "" or "").strip()

        g = str(cvals[6] if len(cvals) >= 7 else "" or "").strip()

        h_raw = cvals[7] if len(cvals) >= 8 else None

        h = str(h_raw).strip() if h_raw is not None and str(h_raw).strip() != "" else ""

        denom = d.split("/", 1)[1] if "/" in d else ""

        if g and "/" not in g and denom:

            g = f"{g}/{denom}"

        if c and c not in matched_w_docs:

            matched_w_docs.append(c)

        if e and e.upper() not in {"NA", "N/A", "-"} and e not in matched_w_revs:

            matched_w_revs.append(e)

        if d and d not in matched_w_sheets:

            matched_w_sheets.append(d)

        if f and f not in matched_d_docs:

            matched_d_docs.append(f)

        if h and h.upper() not in {"NA", "N/A", "-"} and h not in matched_d_revs:

            matched_d_revs.append(h)

        if g and g not in matched_d_sheets:

            matched_d_sheets.append(g)

    clean_target = rfi_id.strip().lower()

    if report_type == "ICR" and pipe_table:

        candidate_cols = [17, 16]

        if len(pipe_table) > 0 and pipe_table[0][0] == 1:

            hdr = pipe_table[0][1]

            found = [idx + 1 for idx, h in enumerate(hdr) if "ICR" in str(h or "").upper() and "LINK" not in str(h or "").upper()]

            if found:

                candidate_cols = found

        pipe_numbers = set()

        for r_num, pvals in pipe_table:

            if r_num < 2:

                continue

            is_match = False

            for c_idx in candidate_cols:

                icr_val = str(pvals[c_idx - 1] or "").strip().lower() if len(pvals) >= c_idx else ""

                if icr_val == clean_target:

                    is_match = True

                    break

            pipe_val = str(pvals[11] or "").strip() if len(pvals) >= 12 else ""

            if is_match and pipe_val:

                pipe_numbers.add(pipe_val)

        if pipe_numbers:

            for r_num, cvals in control_table:

                if r_num < 4:

                    continue

                pipe1 = str(cvals[17] or "").strip() if len(cvals) >= 18 else ""

                pipe2 = str(cvals[24] or "").strip() if len(cvals) >= 25 else ""

                if pipe1 in pipe_numbers or pipe2 in pipe_numbers:

                    add_row_match(cvals)

    else:

        for row_num, values in control_table:

            if row_num < 4:

                continue

            is_match = False

            if col_idx is not None and len(values) >= col_idx:

                val_str = str(values[col_idx - 1] or "").strip().lower()

                if val_str == clean_target:

                    is_match = True

                elif len(values) >= col_idx + 2 and str(values[col_idx + 1] or "").strip().lower() == clean_target:

                    is_match = True

            elif col_idx is None:

                # No dedicated assignment column — try exact match anywhere in row

                for v in values:

                    v_str = str(v or "").strip().lower()

                    if v_str and v_str == clean_target:

                        is_match = True

                        break

            if is_match:

                add_row_match(values)

    if matched_w_docs:

        w_doc = re.sub(r"\s+Rev.*$", "", matched_w_docs[0], flags=re.IGNORECASE).strip()

        w_parts = [w_doc]

        if matched_w_revs:

            w_parts.append(f"REV.{matched_w_revs[0]}")

        formatted_w_sheet = format_sheet_numbers(matched_w_sheets)

        if formatted_w_sheet:

            w_parts.append(f"SHT {formatted_w_sheet}")

        res_weld_map = " ".join(w_parts)

    else:

        res_weld_map = fallback_weld_map

    if matched_d_docs:

        d_doc = re.sub(r"\s+Rev.*$", "", matched_d_docs[0], flags=re.IGNORECASE).strip()

        d_parts = [d_doc]

        if matched_d_revs:

            d_parts.append(f"REV.{matched_d_revs[0]}")

        formatted_d_sheet = format_sheet_numbers(matched_d_sheets)

        if formatted_d_sheet:

            d_parts.append(f"SHT {formatted_d_sheet}")

        res_drawing = " ".join(d_parts)

    else:

        res_drawing = fallback_drawing

    return res_drawing, res_weld_map

def resolve_weld_map_for_rfi(

    rfi_id: str,

    report_type: str,

    external_tables: dict[str, list[tuple[int, tuple[object, ...]]]],

    location: str = DEFAULT_SITE,

    fallback_weld_map: str = DEFAULT_WELD_MAP_NO,

) -> str:

    """Tương thích ngược: lấy số bản vẽ mối hàn từ resolve_drawings_for_rfi."""

    _, wmp = resolve_drawings_for_rfi(

        rfi_id, report_type, external_tables, location, fallback_weld_map=fallback_weld_map

    )

    return wmp

def detect_rp_layout(ws) -> dict[str, object] | None:

    page_cell = None

    for r in range(1, 11):

        for c in range(1, 16):

            v = str(ws.cell(row=r, column=c).value or "")

            if "Trang" in v or "Page" in v:

                page_cell = (r, c)

                break

        if page_cell:

            break

    hdr_row = None

    for r in range(1, 20):

        c1 = str(ws.cell(row=r, column=1).value or "").upper()

        if (

            ("STT" in c1 or "INSTRUMENT" in c1 or "LOẠI" in c1)

            and "BẢN VẼ" not in c1

            and "DRAWING" not in c1

            and "FORM" not in c1

            and "NGÀY" not in c1

        ):

            hdr_row = r

            break

    if not hdr_row:

        return None

    # Check if hdr_row in col 1 is vertically merged (e.g. A9:A10 in MTR)

    table_start = hdr_row + 1

    for m in ws.merged_cells.ranges:

        if m.min_col == 1 and m.min_row <= hdr_row <= m.max_row:

            table_start = m.max_row + 1

            break

    table_end = table_start

    ftr_row = None

    footer_keywords = [

        "GHI CHÚ", "NOTE", "KẾT LUẬN", "CONCLUDE", "KIỂM TRA", "INSPECTION",

        "VIẾT TẮT", "ABBREVIATION", "ĐẠI DIỆN", "REPRESENTATIVES",

        "TIÊU CHÍ", "CRITERIA", "ACCEPTANCE",

    ]

    for r in range(table_start, min(table_start + 40, ws.max_row + 1)):

        c1 = str(ws.cell(row=r, column=1).value or "").upper()

        if any(kw in c1 for kw in footer_keywords):

            ftr_row = r

            table_end = r - 1

            break

    if not ftr_row:

        return None

    sig_end_row = ftr_row

    note_keywords = ["CHÚ THÍCH", "ACCEPTED", "REJECTED", "NOT APPLICABLE", "ACC =", "REJ ="]

    for r in range(ftr_row, min(ftr_row + 20, ws.max_row + 1)):

        row_str = " ".join(str(ws.cell(row=r, column=c).value or "") for c in range(1, ws.max_column + 1))

        if any(kw in row_str.upper() for kw in note_keywords):

            sig_end_row = r

            break

    next_row_empty = True

    for c in range(1, ws.max_column + 1):

        if ws.cell(row=sig_end_row + 1, column=c).value is not None:

            next_row_empty = False

            break

    page_len = sig_end_row + 1 if next_row_empty else sig_end_row

    key_col = None

    for c in range(1, ws.max_column + 1):

        cell = ws.cell(row=table_start, column=c)

        if cell.data_type == "f" and "MATCH" in str(cell.value).upper() and "|" in str(cell.value):

            key_col = c

            break

    if not key_col:

        for cand in [2, 4, 1]:

            if cand <= ws.max_column and ws.cell(row=table_start, column=cand).value is not None:

                key_col = cand

                break

    last_col = 1

    for c in range(1, ws.max_column + 1):

        cell = ws.cell(row=table_start, column=c)

        hdr = ws.cell(row=hdr_row, column=c)

        if (cell.border and cell.border.right and cell.border.right.style) or hdr.value is not None:

            last_col = max(last_col, c)

    for m in ws.merged_cells.ranges:

        if m.min_row <= table_start <= m.max_row and m.min_col <= last_col <= m.max_col:

            last_col = max(last_col, m.max_col)

    is_stt = "STT" in str(ws.cell(row=hdr_row, column=1).value or "").upper()

    return {

        "page_cell": page_cell,

        "hdr_row": hdr_row,

        "table_start": table_start,

        "table_end": table_end,

        "capacity": table_end - table_start + 1,

        "page_len": page_len,

        "key_col": key_col,

        "last_col": last_col,

        "is_stt": is_stt,

    }

def get_rfi_report_items(

    report_type: str,

    rfi_id: str,

    control_table: list[tuple[int, tuple[object, ...]]],

    pipe_table: list[tuple[int, tuple[object, ...]]],

) -> list[object]:

    target_id = rfi_id.strip()

    if report_type == "ICR":

        candidate_cols = [17, 16]

        if pipe_table and len(pipe_table) > 0 and pipe_table[0][0] == 1:

            hdr = pipe_table[0][1]

            found = [idx + 1 for idx, h in enumerate(hdr) if "ICR" in str(h or "").upper() and "LINK" not in str(h or "").upper()]

            if found:

                candidate_cols = found

        val_idx = column_index_from_string("L")

        items = []

        for row_num, vals in pipe_table:

            if row_num < 2:

                continue

            is_match = False

            for c_idx in candidate_cols:

                rfi_val = vals[c_idx - 1] if len(vals) >= c_idx else None

                if rfi_val and str(rfi_val).strip() == target_id:

                    is_match = True

                    break

            if is_match:

                item_val = vals[val_idx - 1] if len(vals) >= val_idx else None

                if item_val:

                    items.append(item_val)

        return items

    elif report_type == "MRIR":

        col_idx = column_index_from_string("O")

        val_idx = column_index_from_string("L")

        items = []

        for row_num, vals in pipe_table:

            if row_num < 2:

                continue

            rfi_val = vals[col_idx - 1] if len(vals) >= col_idx else None

            item_val = vals[val_idx - 1] if len(vals) >= val_idx else None

            if rfi_val and str(rfi_val).strip() == target_id and item_val:

                items.append(item_val)

        return items

    elif report_type in RP_ASSIGNMENT_COLUMNS:

        col_letter = RP_ASSIGNMENT_COLUMNS[report_type]

        col_idx = column_index_from_string(col_letter)

        joint_idx = column_index_from_string("I")

        items = []

        for row_num, vals in control_table:

            if row_num < 4:

                continue

            rfi_val = vals[col_idx - 1] if len(vals) >= col_idx else None

            joint_val = vals[joint_idx - 1] if len(vals) >= joint_idx else None

            if rfi_val and str(rfi_val).strip() == target_id and joint_val:

                items.append(joint_val)

        return items

    return []

def update_drawing_anchors_and_locks(

    xml_bytes: bytes,

    num_pages: int,

    page_len: int,

    empty_row_ranges: list[tuple[int, int, int, int]] | None = None,

) -> bytes:

    ElementTree.register_namespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")

    ElementTree.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")

    ElementTree.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    ElementTree.register_namespace("a14", "http://schemas.microsoft.com/office/drawing/2010/main")

    ElementTree.register_namespace("a16", "http://schemas.microsoft.com/office/drawing/2014/main")

    try:

        root = ElementTree.fromstring(xml_bytes)

    except Exception:

        return xml_bytes

    if num_pages > 1:

        base_anchors = list(root)

        for p in range(2, num_pages + 1):

            offset = page_len * (p - 1)

            for anchor in base_anchors:

                new_anchor = copy.deepcopy(anchor)

                f = new_anchor.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}from")

                if f is not None:

                    r = f.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}row")

                    if r is not None and r.text:

                        r.text = str(int(r.text) + offset)

                t = new_anchor.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}to")

                if t is not None:

                    r = t.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}row")

                    if r is not None and r.text:

                        r.text = str(int(r.text) + offset)

                for cnv in new_anchor.iter("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}cNvPr"):

                    if "id" in cnv.attrib:

                        try:

                            cnv.attrib["id"] = str(int(cnv.attrib["id"]) + 100 * p)

                        except ValueError:

                            pass

                root.append(new_anchor)

    xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"

    # Remove existing connector shapes to prevent duplicate/lingering diagonal lines

    for child in list(root):

        is_connector = False

        if child.find(f".//{{{xdr_ns}}}cxnSp") is not None:

            is_connector = True

        else:

            geom = child.find(f".//{{{a_ns}}}prstGeom")

            if geom is not None and geom.attrib.get("prst") == "line":

                is_connector = True

            elif child.find(f".//{{{xdr_ns}}}nvCxnSpPr") is not None:

                is_connector = True

        if is_connector:

            root.remove(child)

    if empty_row_ranges:

        max_id = 100

        for cnv in root.iter(f"{{{xdr_ns}}}cNvPr"):

            if "id" in cnv.attrib:

                try:

                    max_id = max(max_id, int(cnv.attrib["id"]))

                except ValueError:

                    pass

        for from_col, from_row, to_col, to_row in empty_row_ranges:

            if from_row >= to_row or from_col >= to_col:

                continue

            max_id += 1

            anchor = ElementTree.SubElement(root, f"{{{xdr_ns}}}twoCellAnchor")

            frm = ElementTree.SubElement(anchor, f"{{{xdr_ns}}}from")

            c1 = ElementTree.SubElement(frm, f"{{{xdr_ns}}}col"); c1.text = str(from_col)

            co1 = ElementTree.SubElement(frm, f"{{{xdr_ns}}}colOff"); co1.text = "0"

            r1 = ElementTree.SubElement(frm, f"{{{xdr_ns}}}row"); r1.text = str(from_row)

            ro1 = ElementTree.SubElement(frm, f"{{{xdr_ns}}}rowOff"); ro1.text = "0"

            to_el = ElementTree.SubElement(anchor, f"{{{xdr_ns}}}to")

            c2 = ElementTree.SubElement(to_el, f"{{{xdr_ns}}}col"); c2.text = str(to_col)

            co2 = ElementTree.SubElement(to_el, f"{{{xdr_ns}}}colOff"); co2.text = "0"

            r2 = ElementTree.SubElement(to_el, f"{{{xdr_ns}}}row"); r2.text = str(to_row)

            ro2 = ElementTree.SubElement(to_el, f"{{{xdr_ns}}}rowOff"); ro2.text = "0"

            cxn = ElementTree.SubElement(anchor, f"{{{xdr_ns}}}cxnSp")

            nv = ElementTree.SubElement(cxn, f"{{{xdr_ns}}}nvCxnSpPr")

            ElementTree.SubElement(nv, f"{{{xdr_ns}}}cNvPr", attrib={"id": str(max_id), "name": f"Straight Connector {max_id}"})

            ElementTree.SubElement(nv, f"{{{xdr_ns}}}cNvCxnSpPr")

            spPr = ElementTree.SubElement(cxn, f"{{{xdr_ns}}}spPr")

            xfrm = ElementTree.SubElement(spPr, f"{{{a_ns}}}xfrm")

            ElementTree.SubElement(xfrm, f"{{{a_ns}}}off", attrib={"x": "0", "y": "0"})

            ElementTree.SubElement(xfrm, f"{{{a_ns}}}ext", attrib={"cx": "0", "cy": "0"})

            geom = ElementTree.SubElement(spPr, f"{{{a_ns}}}prstGeom", attrib={"prst": "line"})

            ElementTree.SubElement(geom, f"{{{a_ns}}}avLst")

            ln = ElementTree.SubElement(spPr, f"{{{a_ns}}}ln", attrib={"w": "12700"})

            sf = ElementTree.SubElement(ln, f"{{{a_ns}}}solidFill")

            ElementTree.SubElement(sf, f"{{{a_ns}}}srgbClr", attrib={"val": "000000"})

            ElementTree.SubElement(anchor, f"{{{xdr_ns}}}clientData")

    return ElementTree.tostring(root, encoding="utf-8")

def remove_rp_header_dividers(sheet, num_pages: int = 1, page_len: int = 0) -> None:

    """Xóa bỏ gạch ngăn cách giữa nhãn và nội dung cho dòng Khu vực/Area và Hạng mục/Item trên tất cả các biểu mẫu RP."""

    def find_merged_range(ws, row, col):

        for m in ws.merged_cells.ranges:

            if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:

                return m

        return None

    def remove_v_border_between(ws, min_col_l: int, max_col_l: int, min_col_r: int, max_col_r: int, r_start: int, r_end: int):

        for r in range(r_start, r_end + 1):

            for c in range(min_col_l, max_col_l + 1):

                cell_l = ws.cell(row=r, column=c)

                b_l = cell_l.border

                if b_l is not None:

                    cell_l.border = Border(

                        left=b_l.left,

                        top=b_l.top,

                        bottom=b_l.bottom,

                        right=None,

                        diagonal=b_l.diagonal,

                        diagonal_direction=b_l.diagonal_direction,

                    )

            for c in range(min_col_r, max_col_r + 1):

                cell_r = ws.cell(row=r, column=c)

                b_r = cell_r.border

                if b_r is not None:

                    cell_r.border = Border(

                        right=b_r.right,

                        top=b_r.top,

                        bottom=b_r.bottom,

                        left=None,

                        diagonal=b_r.diagonal,

                        diagonal_direction=b_r.diagonal_direction,

                    )

    found_kv = None

    found_hm = None

    limit_c = min(sheet.max_column, 32)

    for r in range(5, 11):

        for c in range(1, limit_c + 1):

            val = sheet.cell(row=r, column=c).value

            if val is None:

                continue

            s = str(val).strip().lower()

            if ("địa điểm" in s or "location" in s or "khu vực" in s or "area" in s) and not found_kv:

                found_kv = (r, c)

            if (

                ("hạng mục" in s or "item" in s)

                and not found_hm

                and "mã" not in s

                and "loại" not in s

                and "mục tiêu" not in s

            ):

                found_hm = (r, c)

    kv_coords = []

    if found_kv:

        r, c = found_kv

        m_l = find_merged_range(sheet, r, c)

        min_l = m_l.min_col if m_l else c

        max_l = m_l.max_col if m_l else c

        r_start = m_l.min_row if m_l else r

        r_end = m_l.max_row if m_l else r

        m_r = find_merged_range(sheet, r, max_l + 1)

        min_r = m_r.min_col if m_r else (max_l + 1)

        max_r = m_r.max_col if m_r else (max_l + 1)

        kv_coords.append((min_l, max_l, min_r, max_r, r_start, r_end))

    hm_coords = []

    if found_hm:

        r, c = found_hm

        m_l = find_merged_range(sheet, r, c)

        min_l = m_l.min_col if m_l else c

        max_l = m_l.max_col if m_l else c

        r_start = m_l.min_row if m_l else r

        r_end = m_l.max_row if m_l else r

        m_r = find_merged_range(sheet, r, max_l + 1)

        min_r = m_r.min_col if m_r else (max_l + 1)

        max_r = m_r.max_col if m_r else (max_l + 1)

        hm_coords.append((min_l, max_l, min_r, max_r, r_start, r_end))

    pages_to_process = max(1, num_pages)

    for p in range(pages_to_process):

        p_offset = page_len * p if page_len > 0 else 0

        for min_l, max_l, min_r, max_r, r_s, r_e in kv_coords:

            remove_v_border_between(sheet, min_l, max_l, min_r, max_r, r_s + p_offset, r_e + p_offset)

        for min_l, max_l, min_r, max_r, r_s, r_e in hm_coords:

            remove_v_border_between(sheet, min_l, max_l, min_r, max_r, r_s + p_offset, r_e + p_offset)

_FONT_CACHE: dict[tuple, Font] = {}

def _get_black_font(base_font=None, bold=None, italic=None) -> Font:

    f_name = (base_font.name if base_font else None) or "Times New Roman"

    f_size = (base_font.size if base_font else None) or 10

    f_bold = (base_font.bold if base_font else False) if bold is None else bold

    f_italic = (base_font.italic if base_font else False) if italic is None else italic

    key = (f_name, f_size, bool(f_bold), bool(f_italic))

    cached = _FONT_CACHE.get(key)

    if cached is None:

        cached = Font(

            name=f_name,

            size=f_size,

            bold=f_bold,

            italic=f_italic,

            color="000000",

        )

        _FONT_CACHE[key] = cached

    return cached

def normalize_sheet_fonts_and_signature(ws, engineer_name: str = "", max_row_limit: int = 80) -> None:

    """Đảm bảo toàn bộ chữ có màu đen mặc định (000000) và chuẩn hóa in đậm/nhạt

    phần ký Đại diện các bên theo đúng hình ảnh đính kèm:

    - Mọi ô có font: màu đen mặc định (color="000000")

    - Đại diện các bên/Representatives of the parties: in đậm (bold=True)

    - TỔNG THẦU EPC..., GIÁM SÁT THI CÔNG..., CHỦ ĐẦU TƯ..., TƯ VẤN ĐĂNG KIỂM...: in đậm (bold=True)

    - Kỹ thuật thi công..., Giám sát thi công..., Đại diện Chủ đầu tư...: in đậm (bold=True)

    - Kiểm tra/Inspection, Chức danh/Position: không in đậm (bold=False)

    - Chữ ký/Signature, Họ và tên/Full name: không in đậm, in nghiêng (bold=False, italic=True)

    - Tên người ký (trong hàng Họ và tên): không in đậm (bold=False, italic=False)

    - Tự động thay thế tên kỹ sư theo vị trí (engineer_name)

    """

    if ws.title.casefold().startswith("controlog") or ws.title == "MRIR":

        return

    limit_r = ws.max_row if max_row_limit is None else min(ws.max_row, max_row_limit)

    limit_c = min(ws.max_column, 32)

    for (r, c), cell in list(ws._cells.items()):
        if r > limit_r or c > limit_c or r < 12:
            continue
        if isinstance(cell.value, CellRichText):
            continue
        val = str(cell.value or "").strip()
        if not val:
            if cell.font:
                cell.font = _get_black_font(cell.font)
            continue

        if "Đại diện các bên" in val or "Representatives of the parties" in val:
            cell.font = _get_black_font(cell.font, bold=True, italic=False)
        elif r >= 10 and any(k in val for k in ["TỔNG THẦU", "GIÁM SÁT THI CÔNG", "CHỦ ĐẦU TƯ", "TƯ VẤN ĐĂNG KIỂM"]):
            cell.font = _get_black_font(cell.font, bold=True, italic=False)
        elif any(k in val for k in ["Kỹ thuật thi công", "Giám sát thi công", "Đại diện Chủ đầu tư", "Đại diện chủ đầu tư", "Company Representative"]):
            cell.font = _get_black_font(cell.font, bold=True, italic=False)
        elif any(k in val for k in ["Kiểm tra/", "Kiểm Tra/", "Chức danh/", "Chữ ký", "Họ và tên"]):
            cell.font = _get_black_font(cell.font, bold=False, italic=False)
        elif any(nm in val for nm in ["Nguyễn Đức", "Duc Son", "Lê Đình Trường", "Lê Quốc Đoàn", "Võ Xuân Duy", "Nguyễn Lương Định", "Quốc Đoàn", "Xuân Duy"]):
            cell.font = _get_black_font(cell.font, bold=False, italic=False)

def format_ppp_sheet(

    ws,

    d_str: str,

    rep_no: str,

    rfi_id: str,

    drawing_no: str,

    weld_map_no: str,

    location_str: str,

    engineer_name: str,

    num_pages: int = 1,

    page_len: int = 22,

) -> None:

    """Format PPP report worksheet to match reference image 100%.

    Applies exact rich text, bold/regular, upright/italic styles, and exact line breaks (xuống dòng).

    """

    f10_b = InlineFont(rFont="Times New Roman", sz=10, b=True)

    f10_i = InlineFont(rFont="Times New Roman", sz=10, i=True)

    f10_r = InlineFont(rFont="Times New Roman", sz=10)

    sz_body = 10
    f9_b = InlineFont(rFont="Times New Roman", sz=10, b=True)

    f9_i = InlineFont(rFont="Times New Roman", sz=10, i=True)

    f9_r = InlineFont(rFont="Times New Roman", sz=10)
    f8_b = InlineFont(rFont="Times New Roman", sz=sz_body, b=True)
    f8_i = InlineFont(rFont="Times New Roman", sz=sz_body, i=True)
    f8_r = InlineFont(rFont="Times New Roman", sz=sz_body)
    f8_bi = InlineFont(rFont="Times New Roman", sz=sz_body, b=True, i=True)

    f12_b = InlineFont(rFont="Times New Roman", sz=12, b=True)

    f12_i = InlineFont(rFont="Times New Roman", sz=12, i=True)

    f12_r = InlineFont(rFont="Times New Roman", sz=12)
    f14_b = InlineFont(rFont="Times New Roman", sz=14, b=True)
    f14_r = InlineFont(rFont="Times New Roman", sz=14, b=False, i=False)
    font_data_10 = Font(name="Times New Roman", size=10, bold=False, italic=False, color="000000")

    font_bold_10 = Font(name="Times New Roman", size=10, bold=True)

    font_reg_10 = Font(name="Times New Roman", size=10, bold=False, italic=False)

    font_data_9 = Font(name="Times New Roman", size=9, bold=False, italic=False, color="000000")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    clean_draw = str(drawing_no or "").strip()

    if ": " in clean_draw:

        clean_draw = clean_draw.split(": ", 1)[1].strip()

    elif ":" in clean_draw:

        clean_draw = clean_draw.split(":", 1)[1].strip()

    if not clean_draw:

        cur_a7 = str(ws.cell(row=7, column=1).value or "").strip()

        if ": " in cur_a7:

            clean_draw = cur_a7.split(": ", 1)[1].strip()

    clean_wmp = str(weld_map_no or "").strip()

    if ": " in clean_wmp:

        clean_wmp = clean_wmp.split(": ", 1)[1].strip()

    elif ":" in clean_wmp:

        clean_wmp = clean_wmp.split(":", 1)[1].strip()

    if not clean_wmp:

        cur_f7 = str(ws.cell(row=7, column=6).value or "").strip()

        if ": " in cur_f7:

            clean_wmp = cur_f7.split(": ", 1)[1].strip()

    for p in range(1, num_pages + 1):

        offset = (p - 1) * page_len

        # Row 1: Biểu mẫu số/Form No: SWG-70-PL-7001-QCF-008

        c_a1 = ws.cell(row=1 + offset, column=1)

        c_a1.value = CellRichText(

            TextBlock(f10_r, "Biểu mẫu số/"),

            TextBlock(f10_i, "Form No: "),

            TextBlock(f10_r, "SWG-70-PL-7001-QCF-008"),

        )

        c_a1.alignment = align_right

        # Row 2: C2: ĐƯỜNG ỐNG DẪN KHÍ LÔ B - Ô MÔN / BLOCK B – O MON GAS PIPELINE PROJECT (font 12)

        c_c2 = ws.cell(row=2 + offset, column=3)

        c_c2.value = CellRichText(

            TextBlock(f14_b, "DỰ ÁN ĐƯỜNG ỐNG DẪN KHÍ LÔ B- Ô MÔN/\n"),

            TextBlock(f14_r, "BLOCK B – O MON GAS PIPELINE PROJECT"),

        )

        c_c2.font = Font(name="Times New Roman", size=14, bold=False, italic=False)

        c_c2.alignment = align_center

        # Row 3: C3: BÁO CÁO KIỂM TRA KÉO ĐẨY ỐNG XUỐNG HÀO NƯỚC / INSPECTION REPORT... (font 12)

        c_c3 = ws.cell(row=3 + offset, column=3)

        c_c3.value = CellRichText(

            TextBlock(f14_b, "BÁO CÁO KIỂM TRA KÉO ĐẨY ỐNG XUỐNG HÀO NƯỚC/\n"),

            TextBlock(f12_i, "INSPECTION REPORT FOR PULLING PUSHING PIPE TO TRENCH"),

        )

        c_c3.font = Font(name="Times New Roman", size=14, bold=True, italic=False)

        c_c3.alignment = align_center

        # Row 3: I3: Ngày/ Date: (both in italic)

        c_i3 = ws.cell(row=3 + offset, column=9)

        c_i3.value = CellRichText(

            TextBlock(f10_i, "Ngày/ Date:"),

        )

        c_i3.alignment = align_left

        # Row 3: J3: Date value

        c_j3 = ws.cell(row=3 + offset, column=10)

        c_j3.value = d_str

        c_j3.font = font_reg_10

        c_j3.alignment = align_left

        # Row 4: I4: Báo cáo/ Report No.:

        c_i4 = ws.cell(row=4 + offset, column=9)

        c_i4.value = CellRichText(

            TextBlock(f10_r, "Báo cáo/ Report No.:"),

        )

        c_i4.alignment = align_left

        # Row 4: J4: Report No. (Bold!)

        c_j4 = ws.cell(row=4 + offset, column=10)

        c_j4.value = rep_no

        c_j4.font = font_reg_10

        c_j4.alignment = align_left

        # Row 5: I5: Trang/ Page:

        c_i5 = ws.cell(row=5 + offset, column=9)

        c_i5.value = CellRichText(

            TextBlock(f10_r, "Trang/ Page:"),

        )

        c_i5.alignment = align_left

        # Row 5: J5: Page number

        c_j5 = ws.cell(row=5 + offset, column=10)

        c_j5.value = str(p)

        c_j5.font = font_reg_10

        c_j5.alignment = align_left

        # Row 6: A6: Địa điểm/ \nLocation: (Bold VN on line 1, Italic EN on line 2)

        c_a6 = ws.cell(row=6 + offset, column=1)

        c_a6.value = CellRichText(

            TextBlock(f10_b, "Địa điểm/\n"),

            TextBlock(f10_i, "Location:"),

        )

        c_a6.alignment = align_left

        # Row 6: B6: Location value (Regular VN on line 1, Regular EN on line 2)

        c_b6 = ws.cell(row=6 + offset, column=2)

        cur_b6 = str(c_b6.value or location_str or "").strip()

        if "/" in cur_b6:

            vn_loc, en_loc = cur_b6.split("/", 1)

            vn_loc = vn_loc.strip() + "/\n"

            en_loc = en_loc.strip()

            c_b6.value = CellRichText(

                TextBlock(f10_r, vn_loc),

                TextBlock(f10_r, en_loc),

            )

        else:

            c_b6.value = CellRichText(TextBlock(f10_r, cur_b6))

        c_b6.alignment = align_left

        # Row 6: E6: Hạng mục /\nItem: (Bold VN on line 1, Italic EN on line 2)

        c_e6 = ws.cell(row=6 + offset, column=5)

        c_e6.value = CellRichText(

            TextBlock(f10_b, "Hạng mục /\n"),

            TextBlock(f10_i, "Item:"),

        )

        c_e6.alignment = align_left

        # Row 6: F6: Thi công tuyến ống bờ/\nOnshore pipeline installation

        c_f6 = ws.cell(row=6 + offset, column=6)

        c_f6.value = CellRichText(

            TextBlock(f10_r, "Thi công tuyến ống bờ/\n"),

            TextBlock(f10_r, "Onshore pipeline installation"),

        )

        c_f6.alignment = align_left

        # Row 6: I6: Thư mời kiểm tra/\nRFI No.: (Bold VN on line 1, Italic EN on line 2)

        c_i6 = ws.cell(row=6 + offset, column=9)

        c_i6.value = CellRichText(

            TextBlock(f10_b, "Thư mời kiểm tra/\n"),

            TextBlock(f10_i, "RFI No.:"),

        )

        c_i6.alignment = align_left

        # Row 6: J6: RFI No. (Bold!)

        c_j6 = ws.cell(row=6 + offset, column=10)

        c_j6.value = rfi_id

        c_j6.font = font_reg_10

        c_j6.alignment = align_left

        # Row 7: A7: Bản vẽ / Drawing No.: <drawing_no>

        c_a7 = ws.cell(row=7 + offset, column=1)

        c_a7.value = CellRichText(

            TextBlock(f10_b, "Bản vẽ / "),

            TextBlock(f10_i, "Drawing No.: "),

            TextBlock(f10_b, clean_draw),

        )

        c_a7.alignment = align_left

        # Row 7: F7: Bản vẽ mối hàn/ Weld map No.: <weld_map_no>

        c_f7 = ws.cell(row=7 + offset, column=6)

        c_f7.value = CellRichText(

            TextBlock(f10_b, "Bản vẽ mối hàn/ "),

            TextBlock(f10_i, "Weld map No.: "),

            TextBlock(f10_b, clean_wmp),

        )

        c_f7.alignment = align_left

        # Rows 8 & 9: Table Headers (Exact line breaks & styles)

        table_headers_spec = {

            (8, 1): [TextBlock(f9_b, "STT/\n"), TextBlock(f9_i, "No.")],

            (8, 2): [TextBlock(f9_b, "Số mối hàn/\n"), TextBlock(f9_i, "Joint No.")],

            (8, 3): [TextBlock(f9_b, "Ống được kéo đẩy xuống hào /\n"), TextBlock(f9_i, "The pipe was pulled & pushed down into the trench.")],

            (8, 5): [TextBlock(f9_b, "Kiểm tra tình trạng hoàn thành NDT, FJC & điền đầy Foam của mối\nhàn/ "), TextBlock(f9_i, "Check the NDT, FJC & infilled Foam completion status of the weld")],

            (8, 8): [TextBlock(f9_b, "Kiểm tra hệ thống\nkéo đẩy ống/ "), TextBlock(f9_i, "Check\nthe pipe push-pull\nsystem")],

            (8, 9): [TextBlock(f9_b, "Kiểm tra cao độ mực\nnước trong hào đủ để\nkéo ống/ "), TextBlock(f9_i, "Check the water\nlevel in the trench is\nenough to pull the pipe")],

            (8, 10): [TextBlock(f9_b, "Kiểm tra tình trạng\nphao/ "), TextBlock(f9_i, "Check the\ncondition of the float")],

            (8, 11): [TextBlock(f9_b, "Ghi chú/\n"), TextBlock(f9_i, "Remark")],

            (9, 3): [TextBlock(f9_b, "Số ống/\n"), TextBlock(f9_i, "Pipe No.")],

            (9, 4): [TextBlock(f9_b, "Chiều dài ống theo MRIR/\n"), TextBlock(f9_i, "Length pipe as per MRIR\n(mm)")],

            (9, 5): [TextBlock(f9_b, "Báo cáo NDT/\n"), TextBlock(f9_i, "NDT report")],

            (9, 6): [TextBlock(f9_b, "Báo cáo FJC/\n"), TextBlock(f9_i, "FJC report")],

            (9, 7): [TextBlock(f9_b, "Báo cáo HDPUF/\n"), TextBlock(f9_i, "HDPUF report")],

        }

        for (r_rel, c_rel), blocks in table_headers_spec.items():

            cell = ws.cell(row=r_rel + offset, column=c_rel)

            cell.value = CellRichText(*blocks)

            cell.alignment = align_center

        # Rows 10..16: Table Data Rows

        for r_rel in range(10, 17):

            for c_rel in range(1, 12):

                cell = ws.cell(row=r_rel + offset, column=c_rel)

                if cell.value is not None and str(cell.value).strip() != "":

                    cell.font = font_data_10

                    cell.alignment = align_center

        # Row 17: Comments & Conclusion (3 lines exact match)

        c_a17 = ws.cell(row=17 + offset, column=1)

        c_a17.value = CellRichText(

            TextBlock(f9_b, "Ý kiến của Chủ đầu tư (nếu có)/"),

            TextBlock(f9_i, "Comment of Company Representative (if any):\n\n"),

            TextBlock(f9_b, "Kết luận/"),

            TextBlock(f9_i, "Conclude:            "),

            TextBlock(f9_r, "☐ ACC - Chấp nhận nghiệm thu, đồng ý cho triển khai công việc tiếp theo/"),

            TextBlock(f9_b, "                                                                                ☐ REJ\n"),

            TextBlock(f9_i, "                                Acceptance of inspection and approval to proceed with the next phase"),

        )

        c_a17.alignment = align_left

        # Row 18: Inspection & Company Headers

        c_a18 = ws.cell(row=18 + offset, column=1)
        c_a18.value = CellRichText(TextBlock(f9_r, "Kiểm tra/\n"), TextBlock(f9_r, "Inspection"))
        c_a18.alignment = align_center

        c_c18 = ws.cell(row=18 + offset, column=3)
        c_c18.value = CellRichText(TextBlock(f10_r, "TỔNG THẦU EPC - PTSCL18/\n"), TextBlock(f10_r, "EPC CONTRACTOR – PTSCL18"))
        c_c18.font = font_reg_10
        c_c18.alignment = align_center

        c_e18 = ws.cell(row=18 + offset, column=5)
        c_e18.value = CellRichText(TextBlock(f10_r, "GIÁM SÁT THI CÔNG IBST/S/\n"), TextBlock(f10_r, "IBST/S CONSTRUCTION\nSUPERVISION"))
        c_e18.font = font_reg_10
        c_e18.alignment = align_center

        c_h18 = ws.cell(row=18 + offset, column=8)
        c_h18.value = CellRichText(TextBlock(f10_b, "CHỦ ĐẦU TƯ-SWPOC/\n"), TextBlock(f10_b, "COMPANY-SWPOC"))
        c_h18.font = font_bold_10
        c_h18.alignment = align_center

        c_j18 = ws.cell(row=18 + offset, column=10)
        c_j18.value = CellRichText(TextBlock(f10_r, "TƯ VẤN ĐĂNG KIỂM/\n"), TextBlock(f10_r, "CA"))
        c_j18.font = font_reg_10
        c_j18.alignment = align_center

        # Row 19: Positions
        c_a19 = ws.cell(row=19 + offset, column=1)
        c_a19.value = CellRichText(TextBlock(f9_r, "Chức danh/\n"), TextBlock(f9_r, "Position"))
        c_a19.alignment = align_center

        c_c19 = ws.cell(row=19 + offset, column=3)
        c_c19.value = CellRichText(TextBlock(f10_r, "Kỹ thuật thi công/\n"), TextBlock(f10_r, "Construction Technician"))
        c_c19.font = font_reg_10
        c_c19.alignment = align_center

        c_e19 = ws.cell(row=19 + offset, column=5)
        c_e19.value = CellRichText(TextBlock(f10_r, "Giám sát thi công/\n"), TextBlock(f10_r, "Construction Supervisor"))
        c_e19.font = font_reg_10
        c_e19.alignment = align_center

        c_h19 = ws.cell(row=19 + offset, column=8)
        c_h19.value = CellRichText(TextBlock(f10_r, "Đại diện Chủ đầu tư/\n"), TextBlock(f10_r, "Company Representative"))
        c_h19.font = font_reg_10
        c_h19.alignment = align_center

        c_j19 = ws.cell(row=19 + offset, column=10)
        c_j19.value = None

        # Row 20: Signatures
        c_a20 = ws.cell(row=20 + offset, column=1)
        c_a20.value = CellRichText(TextBlock(f9_r, "Chữ ký/\n"), TextBlock(f9_r, "Signature"))
        c_a20.alignment = align_center
        c_c20 = ws.cell(row=20 + offset, column=3); c_c20.value = None
        c_e20 = ws.cell(row=20 + offset, column=5); c_e20.value = None
        c_h20 = ws.cell(row=20 + offset, column=8); c_h20.value = None
        c_j20 = ws.cell(row=20 + offset, column=10); c_j20.value = None

        # Row 21: Full name
        c_a21 = ws.cell(row=21 + offset, column=1)
        c_a21.value = CellRichText(TextBlock(f9_r, "Họ và tên/\n"), TextBlock(f9_r, "Full name"))
        c_a21.alignment = align_center
        c_c21 = ws.cell(row=21 + offset, column=3)
        if not c_c21.value:
            c_c21.value = "Nguyễn Đức Sơn"
        c_c21.font = Font(name="Times New Roman", size=10, bold=False, italic=False)
        c_c21.alignment = align_center
        c_e21 = ws.cell(row=21 + offset, column=5); c_e21.value = None
        c_h21 = ws.cell(row=21 + offset, column=8); c_h21.value = None
        c_j21 = ws.cell(row=21 + offset, column=10); c_j21.value = None

        c_a22 = ws.cell(row=22 + offset, column=1)

        c_a22.value = CellRichText(

            TextBlock(f10_r, "    Chú thích/ "),

            TextBlock(f10_i, "Note:                                                "),

            TextBlock(f10_r, "Chấp thuận/ "),

            TextBlock(f10_i, "ACC = Accepted                                                     "),

            TextBlock(f10_r, "Không chấp thuận/ "),

            TextBlock(f10_i, "REJ = Rejected                                                             "),

            TextBlock(f10_r, "Không áp dụng/ "),

            TextBlock(f10_i, "Ne:"),

        )

        c_a22.alignment = align_left

def format_rp_date_str(val: object) -> str:

    if val is None:

        return ""

    if isinstance(val, (datetime, date)):

        return val.strftime("%d/%m/%Y")

    val_str = str(val).strip()

    if not val_str:

        return ""

    m = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", val_str)

    if m:

        return f"{int(m.group(3)):02d}/{int(m.group(2)):02d}/{m.group(1)}"

    m = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})", val_str)

    if m:

        return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}"

    return val_str

def get_icr_pipe_items(

    rfi_id: str,

    pipe_table: list[tuple[int, tuple[object, ...]]],

) -> list[dict[str, str]]:

    target_id = rfi_id.strip()

    target_norm = re.sub(r"[\s\-_]+", "", target_id).upper()

    col_l = 12  # Pipe No

    col_b = 2   # Serial No

    col_d = 4   # Description

    col_h = 8   # OD

    col_mrir = 16 # MRIR Report No

    candidate_icr_cols = []

    if pipe_table and len(pipe_table) > 0 and pipe_table[0][0] == 1:

        hdr = pipe_table[0][1]

        for idx, h in enumerate(hdr):

            h_str = str(h or "").upper()

            if "LINK" in h_str:

                continue

            if "ICR" in h_str:

                if "RFI" in h_str:

                    candidate_icr_cols.insert(0, idx + 1)

                else:

                    candidate_icr_cols.append(idx + 1)

            elif "MRIR" in h_str:

                if "REPORT" in h_str:

                    col_mrir = idx + 1

                elif "RFI" not in h_str and col_mrir == 16:

                    col_mrir = idx + 1

            elif "PIPE NO" in h_str or "SỐ ỐNG" in h_str:

                col_l = idx + 1

            elif "SERINO" in h_str or "SỐ SÊ RI" in h_str or "SERIAL" in h_str:

                col_b = idx + 1

            elif "MÔ TẢ" in h_str or "DESCRIPTION" in h_str:

                col_d = idx + 1

            elif "OD" in h_str or "ĐƯỜNG KÍNH" in h_str:

                col_h = idx + 1

    if not candidate_icr_cols:

        candidate_icr_cols = [17, 16]

    elif 17 not in candidate_icr_cols and 16 not in candidate_icr_cols:

        candidate_icr_cols.extend([17, 16])

    items: list[dict[str, str]] = []

    for row_num, vals in pipe_table:

        if row_num < 2:

            continue

        is_match = False

        for c_icr in candidate_icr_cols:

            if len(vals) >= c_icr:

                rfi_val = vals[c_icr - 1]

                if not rfi_val:

                    continue

                rfi_str = str(rfi_val).strip()

                rfi_norm = re.sub(r"[\s\-_]+", "", rfi_str).upper()

                if rfi_str == target_id or rfi_norm == target_norm or target_id in rfi_str or rfi_str in target_id:

                    is_match = True

                    break

        if is_match:

            pipe_no = str(vals[col_l - 1] if len(vals) >= col_l else "").strip()

            serial_no = str(vals[col_b - 1] if len(vals) >= col_b else "").strip()

            desc = str(vals[col_d - 1] if len(vals) >= col_d else "").strip()

            size_od = str(vals[col_h - 1] if len(vals) >= col_h else "").strip()

            mrir_no = str(vals[col_mrir - 1] if len(vals) >= col_mrir else "").strip()

            if pipe_no and pipe_no not in {"0", "/", "None"}:

                items.append({

                    "pipe_no": pipe_no,

                    "serial_no": serial_no,

                    "desc": desc,

                    "size_od": size_od,

                    "material_spec": "API 5L X65M PSL2\nPSL2 SAWL",

                    "mrir_no": mrir_no,

                    "cleaning_method": "Air Cleaning. Pig",

                    "remarks": "",

                })

    return items

def format_icr_sheet(

    ws,

    d_str: str,

    rep_no: str,

    rfi_id: str,

    drawing_no: str,

    weld_map_no: str,

    location_str: str,

    engineer_name: str,

    pipe_items: list[dict[str, str]],

    num_pages: int = 1,

    page_len: int = 24,

) -> list[tuple[int, int, int, int]]:

    """Format ICR report worksheet to match reference image 100%.

    Applies exact rich text, bold/regular, upright/italic styles, pagination, and diagonal cross slashes.

    """

    f10_b = InlineFont(rFont="Times New Roman", sz=10, b=True)

    f10_i = InlineFont(rFont="Times New Roman", sz=10, i=True)

    f10_r = InlineFont(rFont="Times New Roman", sz=10)

    sz_body = 10
    f9_b = InlineFont(rFont="Times New Roman", sz=10, b=True)

    f9_i = InlineFont(rFont="Times New Roman", sz=10, i=True)

    f9_r = InlineFont(rFont="Times New Roman", sz=10)

    f8_b = InlineFont(rFont="Times New Roman", sz=sz_body, b=True)

    f8_i = InlineFont(rFont="Times New Roman", sz=sz_body, i=True)

    f8_r = InlineFont(rFont="Times New Roman", sz=sz_body)

    font_bold_10 = Font(name="Times New Roman", size=10, bold=True)

    font_reg_10 = Font(name="Times New Roman", size=10, bold=False, italic=False)

    font_bold_9 = Font(name="Times New Roman", size=9, bold=True)

    font_reg_9 = Font(name="Times New Roman", size=9, bold=False, italic=False)

    font_data_8 = Font(name="Times New Roman", size=8, bold=False, italic=False, color="000000")

    font_data_9 = Font(name="Times New Roman", size=9, bold=False, italic=False, color="000000")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    clean_draw = str(drawing_no or "").strip()

    if ": " in clean_draw:

        clean_draw = clean_draw.split(": ", 1)[1].strip()

    elif ":" in clean_draw:

        clean_draw = clean_draw.split(":", 1)[1].strip()

    clean_draw = re.sub(r"\s+(?:SHT|Sheet).*$", "", clean_draw, flags=re.IGNORECASE).strip()

    clean_wmp = str(weld_map_no or "").strip()

    if ": " in clean_wmp:

        clean_wmp = clean_wmp.split(": ", 1)[1].strip()

    elif ":" in clean_wmp:

        clean_wmp = clean_wmp.split(":", 1)[1].strip()

    clean_wmp = re.sub(r"\s+(?:SHT|Sheet).*$", "", clean_wmp, flags=re.IGNORECASE).strip()

    station_code = station_code_for_location(location_str)

    # 1. Duplicate pages if num_pages > 1

    if num_pages > 1:

        for p in range(2, num_pages + 1):

            offset = page_len * (p - 1)

            for r in range(1, page_len + 1):

                if r in ws.row_dimensions:

                    ws.row_dimensions[r + offset].height = ws.row_dimensions[r].height

                for c in range(1, 13):

                    src_cell = ws.cell(row=r, column=c)

                    dst_cell = ws.cell(row=r + offset, column=c)

                    dst_cell.value = src_cell.value

                    dst_cell.data_type = src_cell.data_type

                    if src_cell.has_style:

                        dst_cell.font = copy.copy(src_cell.font)

                        dst_cell.border = copy.copy(src_cell.border)

                        dst_cell.fill = copy.copy(src_cell.fill)

                        dst_cell.number_format = src_cell.number_format

                        dst_cell.protection = copy.copy(src_cell.protection)

                        dst_cell.alignment = copy.copy(src_cell.alignment)

            for merge in list(ws.merged_cells.ranges):

                if merge.min_row >= 1 and merge.max_row <= page_len:

                    new_merge = f"{get_column_letter(merge.min_col)}{merge.min_row + offset}:{get_column_letter(merge.max_col)}{merge.max_row + offset}"

                    ws.merge_cells(new_merge)

    if num_pages > 1:

        ws.print_area = f"A1:L{page_len * num_pages}"

    else:

        ws.print_area = None

    empty_row_ranges: list[tuple[int, int, int, int]] = []

    # Get supervisor info from template cell E19 and E20 if present

    cur_e19 = str(ws.cell(row=19, column=5).value or "").strip()

    if "\n" in cur_e19:

        sup_vn, sup_en = cur_e19.split("\n", 1)

        sup_vn = sup_vn.strip() + "\n"

        sup_en = sup_en.strip()

    elif "/" in cur_e19:

        sup_vn, sup_en = cur_e19.split("/", 1)

        sup_vn = sup_vn.strip() + "/\n"

        sup_en = sup_en.strip()

    else:

        sup_vn = "GIÁM SÁT THI CÔNG IBST/S\n"

        sup_en = "IBST/S CONSTRUCTION SUPERVISOR"

    cur_e20 = str(ws.cell(row=20, column=5).value or "").strip()

    if "\n" in cur_e20:

        pos_vn, pos_en = cur_e20.split("\n", 1)

        pos_vn = pos_vn.strip() + "\n"

        pos_en = pos_en.strip()

    elif "/" in cur_e20:

        pos_vn, pos_en = cur_e20.split("/", 1)

        pos_vn = pos_vn.strip() + "/\n"

        pos_en = pos_en.strip()

    else:

        pos_vn = "Giám sát thi công/\n"

        pos_en = "Construction Supervisor"

    # Format all pages

    for p in range(1, num_pages + 1):

        offset = (p - 1) * page_len

        # Keep template signature row heights so page fits exactly 1 page

        pass

        # Row 1: Biểu mẫu số/Form No: SWG-70-PL-7001-QCF-003

        c_a1 = ws.cell(row=1 + offset, column=1)

        c_a1.value = CellRichText(

            TextBlock(f10_r, "Biểu mẫu số/"),

            TextBlock(f10_i, "Form No: "),

            TextBlock(f10_r, "SWG-70-PL-7001-QCF-003"),

        )

        c_a1.alignment = align_right

        # Row 2: DỰ ÁN ĐƯỜNG ỐNG DẪN KHÍ LÔ B- Ô MÔN / BLOCK B – O MON GAS PIPELINE PROJECT

        c_d2 = ws.cell(row=2 + offset, column=4)

        c_d2.value = CellRichText(

            TextBlock(f10_b, "DỰ ÁN ĐƯỜNG ỐNG DẪN KHÍ LÔ B- Ô MÔN/\n"),

            TextBlock(f10_r, "BLOCK B – O MON GAS PIPELINE PROJECT"),

        )

        c_d2.alignment = align_center

        # Row 3: BÁO CÁO KIỂM TRA LÀM SẠCH BÊN TRONG ỐNG / INTERNAL PIPE CLEANING INSPECTION REPORT

        c_d3 = ws.cell(row=3 + offset, column=4)

        c_d3.value = CellRichText(

            TextBlock(f10_b, "BÁO CÁO KIỂM TRA LÀM SẠCH BÊN TRONG ỐNG/\n"),

            TextBlock(f10_r, "INTERNAL PIPE CLEANING INSPECTION REPORT"),

        )

        c_d3.alignment = align_center

        # Row 3: K3 Date label, L3 Date value

        c_k3 = ws.cell(row=3 + offset, column=11)

        c_k3.value = CellRichText(

            TextBlock(f9_r, "Date/ "),

            TextBlock(f9_i, "Ngày:"),

        )

        c_k3.alignment = align_left

        c_l3 = ws.cell(row=3 + offset, column=12)

        c_l3.value = d_str

        c_l3.font = font_reg_9

        c_l3.alignment = align_left

        # Row 4: K4 Báo cáo/ Report No.:

        c_k4 = ws.cell(row=4 + offset, column=11)

        c_k4.value = CellRichText(

            TextBlock(f9_r, "Báo cáo/ "),

            TextBlock(f9_i, "Report No.:"),

        )

        c_k4.alignment = align_left

        ws.cell(row=4 + offset, column=12).alignment = align_left

        # Row 5: K5 Report No value

        c_k5 = ws.cell(row=5 + offset, column=11)

        c_k5.value = rep_no

        c_k5.font = font_bold_9

        c_k5.alignment = align_left

        ws.cell(row=5 + offset, column=12).alignment = align_left

        # Row 6: K6 Trang/ Page: {p}

        c_k6 = ws.cell(row=6 + offset, column=11)

        c_k6.value = CellRichText(

            TextBlock(f9_r, "Trang/ "),

            TextBlock(f9_i, "Page: "),

            TextBlock(f9_r, str(p)),

        )

        c_k6.alignment = align_left

        ws.cell(row=6 + offset, column=12).alignment = align_left

        # Row 7: Location, Item, RFI No.

        c_a7 = ws.cell(row=7 + offset, column=1)

        c_a7.value = CellRichText(

            TextBlock(f9_b, "Địa điểm/\n"),

            TextBlock(f9_i, "Location:"),

        )

        c_a7.alignment = align_left

        c_c7 = ws.cell(row=7 + offset, column=3)

        c_c7.value = CellRichText(

            TextBlock(f9_r, f"Trạm thi công {station_code} - An Giang/\n"),

            TextBlock(f9_i, f"Construction station {station_code} - An Giang"),

        )

        c_c7.alignment = align_left

        c_e7 = ws.cell(row=7 + offset, column=5)

        c_e7.value = CellRichText(

            TextBlock(f9_b, "Hạng mục / "),

            TextBlock(f9_i, "Item:"),

        )

        c_e7.alignment = align_left

        c_f7 = ws.cell(row=7 + offset, column=6)

        c_f7.value = CellRichText(

            TextBlock(f9_r, "Thi công tuyến ống bờ/\n"),

            TextBlock(f9_i, "Onshore pipeline"),

        )

        c_f7.alignment = align_left

        c_h7 = ws.cell(row=7 + offset, column=8)

        c_h7.value = CellRichText(

            TextBlock(f9_b, "Thư mời kiểm tra/ "),

            TextBlock(f9_i, "RFI No.:"),

        )

        c_h7.alignment = align_left

        c_j7 = ws.cell(row=7 + offset, column=10)

        c_j7.value = rfi_id

        c_j7.font = font_reg_9

        c_j7.alignment = align_center

        # Row 8: Drawings

        c_a8 = ws.cell(row=8 + offset, column=1)

        c_a8.value = CellRichText(

            TextBlock(f9_b, "Bản vẽ / "),

            TextBlock(f9_i, "Drawing No.: "),

            TextBlock(f9_b, clean_draw),

        )

        c_a8.alignment = align_left

        c_g8 = ws.cell(row=8 + offset, column=7)

        c_g8.value = CellRichText(

            TextBlock(f9_b, "Bản vẽ mối hàn/ "),

            TextBlock(f9_i, "Weld map No.: "),

            TextBlock(f9_b, clean_wmp),

        )

        c_g8.alignment = align_left

        # Row 9: Table Headers

        table_headers = {

            1: [TextBlock(f8_b, "STT /\n"), TextBlock(f8_i, "No.")],

            2: [TextBlock(f8_b, "Mô tả /\n"), TextBlock(f8_i, "Description")],

            4: [TextBlock(f8_b, "Kích thước/\n"), TextBlock(f8_i, "Size OD(mm)")],

            5: [TextBlock(f8_b, "Qui cách vật liệu/\n"), TextBlock(f8_i, "Material Specification")],

            6: [TextBlock(f8_b, "Ống số/\n"), TextBlock(f8_i, "Pipe No.")],

            7: [TextBlock(f8_b, "Số nhận dạng/\n"), TextBlock(f8_i, "Serial No.")],

            8: [TextBlock(f8_b, "Báo cáo nhận vật tư/\n"), TextBlock(f8_i, "MRIR Report No.")],

            10: [TextBlock(f8_b, "Phương pháp làm sạch/\n"), TextBlock(f8_i, "Cleaning method")],

            12: [TextBlock(f8_b, "Ghi chú/\n"), TextBlock(f8_i, "Remarks")],

        }

        for col_idx, blocks in table_headers.items():

            cell = ws.cell(row=9 + offset, column=col_idx)

            cell.value = CellRichText(*blocks)

            cell.alignment = align_center

        # Rows 10..17: Table Data Rows

        empty_start = None

        for s in range(8):

            row_num = 10 + offset + s

            item_idx = (p - 1) * 8 + s

            if item_idx < len(pipe_items):

                item = pipe_items[item_idx]

                c_stt = ws.cell(row=row_num, column=1)

                c_stt.value = item_idx + 1

                c_stt.font = font_data_8

                c_stt.alignment = align_center

                c_desc = ws.cell(row=row_num, column=2)

                raw_desc = str(item.get("desc", "")).strip()

                if "3LPE " in raw_desc:

                    raw_desc = raw_desc.replace("3LPE ", "3LPE\n")

                c_desc.value = raw_desc

                c_desc.font = font_data_8

                c_desc.alignment = align_center

                c_size = ws.cell(row=row_num, column=4)

                c_size.value = str(item.get("size_od", "")).strip()

                c_size.font = font_data_8

                c_size.alignment = align_center

                c_mat = ws.cell(row=row_num, column=5)

                c_mat.value = item.get("material_spec", "API 5L X65M PSL2\nPSL2 SAWL")

                c_mat.font = font_data_8

                c_mat.alignment = align_center

                c_pno = ws.cell(row=row_num, column=6)

                c_pno.value = str(item.get("pipe_no", "")).strip()

                c_pno.font = font_data_8

                c_pno.alignment = align_center

                c_sno = ws.cell(row=row_num, column=7)

                c_sno.value = str(item.get("serial_no", "")).strip()

                c_sno.font = font_data_8

                c_sno.alignment = align_center

                c_mrir = ws.cell(row=row_num, column=8)

                c_mrir.value = str(item.get("mrir_no", "")).strip()

                c_mrir.font = font_data_8

                c_mrir.alignment = align_center

                c_clean = ws.cell(row=row_num, column=10)

                c_clean.value = item.get("cleaning_method", "Air Cleaning. Pig")

                c_clean.font = font_data_8

                c_clean.alignment = align_center

                c_rem = ws.cell(row=row_num, column=12)

                c_rem.value = str(item.get("remarks", "")).strip()

                c_rem.font = font_data_8

                c_rem.alignment = align_center

            else:

                if empty_start is None:

                    empty_start = row_num

                for c in range(1, 13):

                    cell = ws.cell(row=row_num, column=c)

                    if not isinstance(cell, MergedCell):

                        cell.value = None

        if empty_start is not None and empty_start <= 17 + offset:

            empty_row_ranges.append((0, empty_start - 1, 12, 17 + offset))

        # Row 18: Comments & Conclusion

        c_a18 = ws.cell(row=18 + offset, column=1)

        c_a18.value = CellRichText(

            TextBlock(f9_b, "Ghi chú /"),

            TextBlock(f9_i, "Note:\n"),

            TextBlock(f9_b, "Kết luận/"),

            TextBlock(f9_i, "Conclude:   "),

            TextBlock(f9_r, "☐ ACC - Chấp nhận nghiệm thu, đồng ý cho triển khai công việc tiếp theo/"),

            TextBlock(f9_b, "               ☐ REJ\n"),

            TextBlock(f9_i, '                                                       Acceptance of inspection and approval to proceed with the next phase"'),

        )

        c_a18.alignment = align_left

        # Row 19: Inspection Headers

        c_a19 = ws.cell(row=19 + offset, column=1)
        c_a19.value = CellRichText(TextBlock(f9_r, "Kiểm tra/\n"), TextBlock(f9_r, "Inspection"))
        c_a19.alignment = align_center

        c_c19 = ws.cell(row=19 + offset, column=3)
        c_c19.value = CellRichText(TextBlock(f9_b, "TỔNG THẦU EPC - PTSCL18/\n"), TextBlock(f9_r, "EPC CONTRACTOR – PTSCL18"))
        c_c19.alignment = align_center

        f8_i = InlineFont(rFont="Times New Roman", sz=8.0, b=False, i=True)
        c_e19 = ws.cell(row=19 + offset, column=5)
        c_e19.value = CellRichText(TextBlock(f9_b, "GIÁM SÁT THI CÔNG IBST/S\n"), TextBlock(f8_i, "IBST/S CONSTRUCTION SUPERVISOR"))
        c_e19.alignment = align_center

        c_h19 = ws.cell(row=19 + offset, column=8)
        c_h19.value = CellRichText(TextBlock(f9_b, "CHỦ ĐẦU TƯ-SWPOC/\n"), TextBlock(f9_r, "COMPANY-SWPOC"))
        c_h19.alignment = align_center

        c_j19 = ws.cell(row=19 + offset, column=10)
        c_j19.value = CellRichText(TextBlock(f9_b, "TƯ VẤN ĐĂNG KIỂM/\n"), TextBlock(f9_r, "CA"))
        c_j19.alignment = align_center

        # Row 20: Positions
        c_a20 = ws.cell(row=20 + offset, column=1)
        c_a20.value = CellRichText(TextBlock(f9_r, "Chức danh/\n"), TextBlock(f9_r, "Position"))
        c_a20.alignment = align_center

        c_c20 = ws.cell(row=20 + offset, column=3)
        c_c20.value = CellRichText(TextBlock(f8_b, "Kỹ thuật thi công/\n"), TextBlock(f8_i, "Construction Technician"))
        c_c20.alignment = align_center

        c_e20 = ws.cell(row=20 + offset, column=5)
        c_e20.value = CellRichText(TextBlock(f8_b, "Giám sát thi công/\n"), TextBlock(f8_i, "Construction Supervisor"))
        c_e20.alignment = align_center

        c_h20 = ws.cell(row=20 + offset, column=8)
        c_h20.value = CellRichText(TextBlock(f8_b, "Đại diện Chủ đầu tư/\n"), TextBlock(f8_i, "Company Representative"))
        c_h20.alignment = align_center

        c_j20 = ws.cell(row=20 + offset, column=10)
        c_j20.value = None

        # Row 21: Signatures
        c_a21 = ws.cell(row=21 + offset, column=1)
        c_a21.value = CellRichText(TextBlock(f9_r, "Chữ ký/\n"), TextBlock(f9_r, "Signature"))
        c_a21.alignment = align_center
        c_c21 = ws.cell(row=21 + offset, column=3); c_c21.value = None
        c_e21 = ws.cell(row=21 + offset, column=5); c_e21.value = None
        c_h21 = ws.cell(row=21 + offset, column=8); c_h21.value = None
        c_j21 = ws.cell(row=21 + offset, column=10); c_j21.value = None

        # Row 22: Full name
        c_a22 = ws.cell(row=22 + offset, column=1)
        c_a22.value = CellRichText(TextBlock(f9_r, "Họ và tên/\n"), TextBlock(f9_r, "Full name"))
        c_a22.alignment = align_center
        c_c22 = ws.cell(row=22 + offset, column=3)
        if not c_c22.value:
            c_c22.value = "Nguyễn Đức Sơn"
        c_c22.font = Font(name="Times New Roman", size=10, bold=False, italic=False)
        c_c22.alignment = align_center
        c_e22 = ws.cell(row=22 + offset, column=5); c_e22.value = None
        c_h22 = ws.cell(row=22 + offset, column=8); c_h22.value = None
        c_j22 = ws.cell(row=22 + offset, column=10); c_j22.value = None
        c_c22.alignment = align_center

        # Row 23: Notes

        c_a23 = ws.cell(row=23 + offset, column=1)

        c_a23.value = CellRichText(

            TextBlock(f9_r, "Chú thích/ "),

            TextBlock(f9_i, "Note:                                 "),

            TextBlock(f9_r, "Chấp thuận/ "),

            TextBlock(f9_i, "ACC = Accepted                            "),

            TextBlock(f9_r, "Không chấp thuận/ "),

            TextBlock(f9_i, "REJ = Rejected                 "),

            TextBlock(f9_r, "Không áp dụng/ "),

            TextBlock(f9_i, "N/A = Not applicable"),

        )

        c_a23.alignment = align_left

    return empty_row_ranges

def format_rp_date_str(d_val) -> str:

    if isinstance(d_val, (datetime, date)):

        return d_val.strftime("%d/%m/%Y")

    elif d_val:

        d_clean = str(d_val).split()[0].replace("-", "/")

        parts = d_clean.split("/")

        if len(parts) == 3 and len(parts[0]) == 4:

            return f"{parts[2]}/{parts[1]}/{parts[0]}"

        return d_clean

    return ""

def get_top_left_cell(ws, row: int, col: int):

    cell = ws.cell(row=row, column=col)

    if isinstance(cell, MergedCell):

        for m in ws.merged_cells.ranges:

            if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:

                return ws.cell(row=m.min_row, column=m.min_col)

    return cell

def format_standard_rp_header(

    ws,

    report_type: str,

    d_str: str,

    rep_no: str,

    rfi_id: str,

    drawing_no: str,

    weld_map_no: str,

    location_str: str,

    engineer_name: str = "Nguyễn Đức Sơn",

    page_num: int = 1,

    num_pages: int = 1,

    offset: int = 0,

) -> None:

    """Format RP report header and info blocks uniformly matching ICR/PPP standard.

    - Project & Report Title: 10pt (portrait) / 12pt (landscape), Line 1 Bold, Line 2 Italic.

    - Top-right box (Date, Report No, Page): Align Left, Middle Align, wrap_text=True, Report No Bold.

    - Info Bar & Drawings Bar: Line 1 Bold, Line 2 Italic.

    - Table column headers: Line 1 Bold, Line 2 Italic, Center horizontal & vertical.

    - Signatures Bar: Line 1 Bold, Line 2 Italic, Engineer name Nguyễn Đức Sơn.

    """

    is_landscape = report_type in {"SPI", "CAI", "FTI", "HTI"}

    is_spi_0001_1 = (report_type == "SPI" and ws.title == "0001-1")

    sz_title = 14 if (report_type == "SPI" or is_spi_0001_1) else (11 if report_type == "FUR" else (12 if is_landscape else 10))

    f_title_b = InlineFont(rFont="Times New Roman", sz=sz_title, b=True)

    f_title_i = InlineFont(rFont="Times New Roman", sz=sz_title, i=True)
    f_title_r = InlineFont(rFont="Times New Roman", sz=sz_title, b=False, i=False)

    f11_b = InlineFont(rFont="Times New Roman", sz=11, b=True)

    f11_i = InlineFont(rFont="Times New Roman", sz=11, i=True)

    f10_b = InlineFont(rFont="Times New Roman", sz=10, b=True)

    f10_i = InlineFont(rFont="Times New Roman", sz=10, i=True)

    f10_r = InlineFont(rFont="Times New Roman", sz=10)

    sz_body = 12 if is_spi_0001_1 else (11 if (report_type in {"DIC", "SPI"}) else 9)

    f9_b = InlineFont(rFont="Times New Roman", sz=sz_body, b=True)

    f9_i = InlineFont(rFont="Times New Roman", sz=sz_body, i=True)

    f9_r = InlineFont(rFont="Times New Roman", sz=sz_body)

    f9_bi = InlineFont(rFont="Times New Roman", sz=sz_body, b=True, i=True)

    f8_b = InlineFont(rFont="Times New Roman", sz=sz_body, b=True)

    f8_i = InlineFont(rFont="Times New Roman", sz=sz_body, i=True)

    f8_r = InlineFont(rFont="Times New Roman", sz=sz_body)

    f8_bi = InlineFont(rFont="Times New Roman", sz=sz_body, b=True, i=True)

    f7_b = InlineFont(rFont="Times New Roman", sz=7, b=True)

    f7_i = InlineFont(rFont="Times New Roman", sz=7, i=True)

    f7_r = InlineFont(rFont="Times New Roman", sz=7)

    f7_bi = InlineFont(rFont="Times New Roman", sz=7, b=True, i=True)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    font_bold_10 = Font(name="Times New Roman", size=10, bold=True)

    font_reg_10 = Font(name="Times New Roman", size=10, bold=False, italic=False)

    font_bold_9 = Font(name="Times New Roman", size=sz_body, bold=True)

    font_reg_9 = Font(name="Times New Roman", size=sz_body, bold=False, italic=False)

    station_code = station_code_for_location(location_str)

    clean_draw = str(drawing_no or "").strip()

    if ": " in clean_draw:

        clean_draw = clean_draw.split(": ", 1)[1].strip()

    elif ":" in clean_draw:

        clean_draw = clean_draw.split(":", 1)[1].strip()

    if report_type in {"ICR", "DIC"}:

        clean_draw = re.sub(r"\s+(?:SHT|Sheet).*$", "", clean_draw, flags=re.IGNORECASE).strip()

    clean_wmp = str(weld_map_no or "").strip()

    if ": " in clean_wmp:

        clean_wmp = clean_wmp.split(": ", 1)[1].strip()

    elif ":" in clean_wmp:

        clean_wmp = clean_wmp.split(":", 1)[1].strip()

    if report_type in {"ICR", "DIC"}:

        clean_wmp = re.sub(r"\s+(?:SHT|Sheet).*$", "", clean_wmp, flags=re.IGNORECASE).strip()

    # 1. Project Title

    proj_coords = {

        "FUR": (2, 3), "VSR": (2, 3), "DIC": (1, 4), "MTR": (2, 3),

        "SPI": (1, 3), "CAI": (1, 3), "FTI": (1, 3), "HTI": (1, 3),

    }

    if report_type in proj_coords:

        r, c = proj_coords[report_type]

        cell = ws.cell(row=r + offset, column=c)

        if not isinstance(cell, MergedCell):

            vn_title = "DỰ ÁN ĐƯỜNG ỐNG DẪN KHÍ LÔ B- Ô MÔN/\n"
            cell.value = CellRichText(
                TextBlock(f_title_b, vn_title),
                TextBlock(f_title_r, "BLOCK B – O MON GAS PIPELINE PROJECT"),
            )

            cell.alignment = align_center

    # 2. Form No in row 1

    form_nos = {

        "FUR": "SWG-70-PL-7001-QCF-005",

        "VSR": "SWG-70-PL-7001-QCF-006",

        "MTR": "SWG-70-PL-7001-QCF-004",

    }

    if report_type in form_nos and report_type != "FUR":

        cell = ws.cell(row=1 + offset, column=1)

        if not isinstance(cell, MergedCell):

            cell.value = CellRichText(

                TextBlock(f10_r, "Biểu mẫu số/"),

                TextBlock(f10_i, "Form No: "),

                TextBlock(f10_r, form_nos[report_type]),

            )

            cell.alignment = align_right

    # 3. Report Title

    rep_titles = {

        "FUR": ("BÁO CÁO KIỂM TRA ĐẤU NỐI MỐI HÀN/\n", "FIT UP INSPECTION REPORT", (3, 3)),

        "VSR": ("BÁO CÁO KIỂM TRA NGOẠI DẠNG MỐI HÀN/\n", "VISUAL INSPECTION REPORT", (3, 3)),

        "DIC": ("BÁO CÁO KIỂM TRA HIỆU CHUẨN THIẾT BỊ HÀNG NGÀY/\n", "DAILY CALIBRATION/ INSPECTION FOR APPARATUS", (4, 4)),

        "MTR": ("BÁO CÁO TRUY SUẤT VẬT TƯ/\n", "MATERIAL TRACEABILITY REPORT", (3, 3)),

        "SPI": ("BÁO CÁO KIỂM TRA CHUẨN BỊ BỀ MẶT/\n" if ws.title != "0001-1" else "DANH MỤC KIỂM TRA TÌNH TRẠNG BỀ MẶT THÉP/\n",

                "SURFACE PREPARATION INSPECTION REPORT" if ws.title != "0001-1" else "CHECKLIST RECORD OF STEEL SURFACE CONDITION",

                (2, 3) if ws.title != "0001-1" else (3, 3)),

        "CAI": ("BÁO CÁO KIỂM TRA THI CÔNG SƠN PHỦ/\n", "COATING APPLICATION INSPECTION REPORT", (3, 3)),

        "FTI": ("BÁO CÁO KIỂM TRA HIỆN TRƯỜNG/\n" if ws.title != "0001-1" else "BÁO CÁO KIỂM TRA SỬA CHỮA MỐI BỌC/\n",

                "FIELD JOINT COATING INSPECTION REPORT" if ws.title != "0001-1" else "COATING JOINT REPAIR INSPECTION REPORT",

                (3, 3)),

        "HTI": ("BÁO CÁO KIỂM TRA HDPUF/\n", "HDPUF TEST INSPECTION REPORT", (3, 3)),

    }

    if report_type in rep_titles:

        vn, en, (r, c) = rep_titles[report_type]

        if report_type == "FUR":

            vn = "BÁO CÁO KIỂM TRA ĐẤU NỐI MỐI HÀN/\n"

        cell = ws.cell(row=r + offset, column=c)

        if not isinstance(cell, MergedCell):

            cell.value = CellRichText(

                TextBlock(f_title_b, vn),

                TextBlock(f_title_r, en),

            )

            cell.alignment = align_center

    # 4. Top-Right Box (Date, Report No, Page)

    top_right_specs = {

        "FUR": {"date_lbl": (3, 13), "date_val": (3, 14), "rep_lbl": (4, 13), "rep_val": (5, 13), "page_lbl": (6, 13)},

        "VSR": {"date_lbl": (3, 13), "date_val": (3, 14), "rep_lbl": (4, 13), "rep_val": (5, 13), "page_lbl": (6, 13)},

        "DIC": {"date_lbl": (3, 15), "date_val": (3, 18), "rep_lbl": (4, 15), "rep_val": (5, 15), "page_lbl": (6, 15)},

        "MTR": {"date_lbl": (3, 11), "date_val": (3, 12), "rep_lbl": (4, 11), "rep_val": (5, 11), "page_lbl": (6, 11)},

        "SPI": {"date_lbl": (2, 9), "date_val": (2, 10), "rep_lbl": (3, 9), "rep_val": (4, 9), "page_lbl": (5, 9)},

        "CAI": {"date_lbl": (2, 9), "date_val": (2, 11), "rep_lbl": (3, 9), "rep_val": (4, 9), "page_lbl": (5, 9)},

        "FTI": {"date_lbl": (2, 10), "date_val": (2, 12), "rep_lbl": (3, 10), "rep_val": (4, 10), "page_lbl": (5, 10)},

        "HTI": {"date_lbl": (2, 10), "date_val": (2, 12), "rep_lbl": (3, 10), "rep_val": (4, 10), "page_lbl": (5, 10)},

    }

    if report_type in top_right_specs:

        spec = dict(top_right_specs[report_type])

        if report_type == "SPI":
            date_col = None
            for c_cand in range(10, 15):
                c_cand_obj = ws.cell(row=2 + offset, column=c_cand)
                if not isinstance(c_cand_obj, MergedCell):
                    date_col = c_cand
                    break
            spec["date_val"] = (2, date_col if date_col else (11 if ws.title == "0001-1" else 10))

        if report_type == "CAI":
            i_j_range = f"I{2 + offset}:J{2 + offset}"
            k_n_range = f"K{2 + offset}:N{2 + offset}"
            existing_merges = {str(m) for m in ws.merged_cells.ranges}
            if i_j_range not in existing_merges:
                ws.merge_cells(start_row=2 + offset, start_column=9, end_row=2 + offset, end_column=10)
            c_j = ws.cell(row=2 + offset, column=10)
            if not isinstance(c_j, MergedCell):
                c_j.value = None
            if k_n_range not in existing_merges:
                ws.merge_cells(start_row=2 + offset, start_column=11, end_row=2 + offset, end_column=14)

        dr, dc = spec["date_lbl"]

        c_lbl = ws.cell(row=dr + offset, column=dc)

        if not isinstance(c_lbl, MergedCell):
            orig_lbl = str(c_lbl.value or "")
            if "ngày" in orig_lbl.lower() and orig_lbl.lower().startswith("ngày"):
                c_lbl.value = CellRichText(TextBlock(f9_r, "Ngày/"), TextBlock(f9_i, "Date:"))
            else:
                c_lbl.value = CellRichText(TextBlock(f9_r, "Date/ "), TextBlock(f9_i, "Ngày:"))

            c_lbl.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False) if report_type == "CAI" else align_left

        dvr, dvc = spec["date_val"]

        c_val = ws.cell(row=dvr + offset, column=dvc)

        if not isinstance(c_val, MergedCell):

            if d_str:

                c_val.value = d_str

            c_val.font = font_reg_9

            c_val.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False) if report_type == "CAI" else align_left

        rr, rc = spec["rep_lbl"]

        c_lbl = ws.cell(row=rr + offset, column=rc)

        if not isinstance(c_lbl, MergedCell):
            orig_rep_lbl = str(c_lbl.value or "")
            if "báo cáo số" in orig_rep_lbl.lower():
                c_lbl.value = CellRichText(TextBlock(f9_r, "Báo cáo số /"), TextBlock(f9_i, "Report No.:"))
            else:
                c_lbl.value = CellRichText(TextBlock(f9_r, "Báo cáo/ "), TextBlock(f9_i, "Report No.:"))

            c_lbl.alignment = align_left

        rvr, rvc = spec["rep_val"]

        c_val = ws.cell(row=rvr + offset, column=rvc)

        if not isinstance(c_val, MergedCell):

            if rep_no:

                c_val.value = rep_no

            c_val.font = font_bold_9

            c_val.alignment = align_left

        pr, pc = spec["page_lbl"]

        c_lbl = ws.cell(row=pr + offset, column=pc)

        if not isinstance(c_lbl, MergedCell):

            p_display = page_num

            if report_type in ["FTI", "SPI"] and ws.title == "0001-1":

                p_display = num_pages + 1

            c_lbl.value = CellRichText(TextBlock(f9_r, "Trang/ "), TextBlock(f9_i, "Page: "), TextBlock(f9_r, str(p_display)))

            c_lbl.alignment = align_left

    # 5. Info Bar (Location, Item, RFI)

    info_specs = {

        "FUR": {"loc_lbl": (7, 1), "loc_val": (7, 3), "item_lbl": (7, 7), "item_val": (7, 8), "rfi_lbl": (7, 11), "rfi_val": (7, 13)},

        "VSR": {"loc_lbl": (7, 1), "loc_val": (7, 2), "item_lbl": (7, 5), "item_val": (7, 6), "rfi_lbl": (7, 9), "rfi_val": (7, 12)},

        "DIC": {"loc_lbl": (7, 1), "loc_val": (7, 4), "item_lbl": (7, 8), "item_val": (7, 9), "rfi_lbl": (7, 15), "rfi_val": (8, 15)},

        "MTR": {"loc_lbl": (7, 1), "loc_val": (7, 2), "item_lbl": (7, 4), "item_val": (7, 5), "rfi_lbl": (7, 8), "rfi_val": (7, 10)},

        "SPI": {"loc_lbl": (6, 1), "loc_val": (6, 2), "item_lbl": (6, 5), "item_val": (6, 6), "rfi_lbl": (6, 8 if ws.title != "0001-1" else 9), "rfi_val": (6 if ws.title != "0001-1" else 7, 10 if ws.title != "0001-1" else 9)},

        "CAI": {"loc_lbl": (6, 1), "loc_val": (6, 2), "item_lbl": (6, 5), "item_val": (6, 6), "rfi_lbl": (6, 8), "rfi_val": (6, 10)},

        "FTI": {"loc_lbl": (6, 1), "loc_val": (6, 2), "item_lbl": (6, 4), "item_val": (6, 5), "rfi_lbl": (6, 7), "rfi_val": (6, 10)},

        "HTI": {"loc_lbl": (6, 1), "loc_val": (6, 2), "item_lbl": (6, 6), "item_val": (6, 7), "rfi_lbl": (6, 10), "rfi_val": (7, 10)},

    }

    if report_type in info_specs:

        s = info_specs[report_type]

        c_lbl = ws.cell(row=s["loc_lbl"][0] + offset, column=s["loc_lbl"][1])

        if not isinstance(c_lbl, MergedCell):

            loc_vn = "Vị trí/\n" if report_type in {"FUR", "MTR"} else "Địa điểm/\n"

            loc_en_font = f9_r if report_type == "FUR" else f9_i

            c_lbl.value = CellRichText(TextBlock(f9_b, loc_vn), TextBlock(loc_en_font, "Location:"))

            c_lbl.alignment = align_center if report_type == "FUR" else align_left

        c_val = ws.cell(row=s["loc_val"][0] + offset, column=s["loc_val"][1])

        if not isinstance(c_val, MergedCell):

            c_val.value = CellRichText(

                TextBlock(f10_r if report_type == "FUR" else f9_r, f"Trạm thi công {station_code} - An Giang/\n"),

                TextBlock(f10_i if report_type == "FUR" else f9_i, f"Construction station {station_code} - An Giang")

            )

            c_val.alignment = align_left

        c_lbl = ws.cell(row=s["item_lbl"][0] + offset, column=s["item_lbl"][1])

        if not isinstance(c_lbl, MergedCell):

            item_vn_font = f9_r if report_type == "FUR" else f9_b

            c_lbl.value = CellRichText(TextBlock(item_vn_font, "Hạng mục / "), TextBlock(f9_i, "Item:"))

            c_lbl.alignment = align_left

        c_val = ws.cell(row=s["item_val"][0] + offset, column=s["item_val"][1])

        if not isinstance(c_val, MergedCell):

            c_val.value = CellRichText(

                TextBlock(f10_r if report_type == "FUR" else f9_r, "Thi công tuyến ống bờ/\n"),

                TextBlock(f10_i if report_type == "FUR" else f9_i, "Onshore pipeline")

            )

            c_val.alignment = align_left

        c_lbl = ws.cell(row=s["rfi_lbl"][0] + offset, column=s["rfi_lbl"][1])

        if not isinstance(c_lbl, MergedCell):

            rfi_vn_font = f9_r if report_type == "FUR" else f9_b

            c_lbl.value = CellRichText(TextBlock(rfi_vn_font, "Thư mời kiểm tra/ \n" if report_type == "FUR" else "Thư mời kiểm tra/ "), TextBlock(f9_i, "RFI No.:"))

            c_lbl.alignment = align_left

        c_val = ws.cell(row=s["rfi_val"][0] + offset, column=s["rfi_val"][1])

        if not isinstance(c_val, MergedCell):

            if rfi_id:

                c_val.value = rfi_id

            c_val.font = font_reg_10 if report_type == "FUR" else font_bold_9

            c_val.alignment = align_left if (is_spi_0001_1 or report_type in {"FUR", "DIC"}) else align_center

    if report_type == "DIC":
        # Ensure O7 (Thư mời kiểm tra/RFI No.) and O8 (RFI No.) are left aligned
        for r_offset in [7, 8]:
            r_idx = r_offset + offset
            for col_idx in range(15, 19):
                ws.cell(row=r_idx, column=col_idx).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        # Remove 2 internal vertical separator lines in rows 7 and 8

        # (between Location label & val: col 3 & 4; and between Item label & val: col 8 & 9)

        for r_offset in [7, 8]:

            r_idx = r_offset + offset

            c3 = ws.cell(r_idx, 3)

            c4 = ws.cell(r_idx, 4)

            if c3.border and c3.border.right:

                c3.border = Border(

                    left=c3.border.left,

                    right=Side(style=None),

                    top=c3.border.top,

                    bottom=c3.border.bottom,

                )

            if c4.border and c4.border.left:

                c4.border = Border(

                    left=Side(style=None),

                    right=c4.border.right,

                    top=c4.border.top,

                    bottom=c4.border.bottom,

                )

            c8 = ws.cell(r_idx, 8)

            c9 = ws.cell(r_idx, 9)

            if c8.border and c8.border.right:

                c8.border = Border(

                    left=c8.border.left,

                    right=Side(style=None),

                    top=c8.border.top,

                    bottom=c8.border.bottom,

                )

            if c9.border and c9.border.left:

                c9.border = Border(

                    left=Side(style=None),

                    right=c9.border.right,

                    top=c9.border.top,

                    bottom=c9.border.bottom,

                )

    if report_type == "FTI":

        # 1. Top-Right Box: remove horizontal divider between row 3 (label) and row 4 (rep_no)

        for c_idx in range(10, 15):

            c3 = ws.cell(row=3 + offset, column=c_idx)

            c4 = ws.cell(row=4 + offset, column=c_idx)

            l3 = Side(style="thin") if c_idx == 10 else (c3.border.left if c3.border else Side(style=None))

            r3 = Side(style="thin") if c_idx == 14 else (c3.border.right if c3.border else Side(style=None))

            t3 = c3.border.top if c3.border and c3.border.top else Side(style="thin")

            c3.border = Border(left=l3, right=r3, top=t3, bottom=Side(style=None))

            l4 = Side(style="thin") if c_idx == 10 else (c4.border.left if c4.border else Side(style=None))

            r4 = Side(style="thin") if c_idx == 14 else (c4.border.right if c4.border else Side(style=None))

            b4 = c4.border.bottom if c4.border and c4.border.bottom else Side(style="thin")

            c4.border = Border(left=l4, right=r4, top=Side(style=None), bottom=b4)

        # 2. Material parameter block in sheet 0001-1 (rows 8 to 11)

        if ws.title == "0001-1":

            thin_s = Side(style="thin")

            no_s = Side(style=None)

            for r_m in range(8, 12):

                r_idx = r_m + offset

                # Fix merged range J:N -> J:O

                for rng in list(ws.merged_cells.ranges):

                    if rng.min_row == r_idx and rng.min_col == 10 and rng.max_col == 14:

                        ws.unmerge_cells(str(rng))

                already_merged_15 = any(

                    rng.min_row == r_idx and rng.min_col == 10 and rng.max_col == 15

                    for rng in ws.merged_cells.ranges

                )

                if not already_merged_15:

                    ws.merge_cells(start_row=r_idx, start_column=10, end_row=r_idx, end_column=15)

                # Left alignment

                c_a = ws.cell(row=r_idx, column=1)

                if not isinstance(c_a, MergedCell):

                    c_a.alignment = align_left

                c_g = ws.cell(row=r_idx, column=7)

                if not isinstance(c_g, MergedCell):

                    c_g.alignment = align_left

                c_j = ws.cell(row=r_idx, column=10)

                if not isinstance(c_j, MergedCell):

                    c_j.alignment = align_left

                # Outer border only (cols 1..15, rows 8..11)

                for col_idx in range(1, 16):

                    cell = ws.cell(row=r_idx, column=col_idx)

                    t = thin_s if r_m == 8 else no_s

                    b = thin_s if r_m == 11 else no_s

                    l = thin_s if col_idx == 1 else no_s

                    r_border = thin_s if col_idx == 15 else no_s

                    cell.border = Border(left=l, right=r_border, top=t, bottom=b)

    if report_type == "SPI" and ws.title != "0001-1":

        # Ensure vertical border between Item (col 7 / G) and RFI (col 8 / H) in row 6

        r_info = 6 + offset

        cg = ws.cell(row=r_info, column=7)

        ch = ws.cell(row=r_info, column=8)

        thin_s = Side(style="thin", color="000000")

        cg.border = Border(

            left=cg.border.left if cg.border else None,

            right=thin_s,

            top=cg.border.top if cg.border else thin_s,

            bottom=cg.border.bottom if cg.border else thin_s,

        )

        ch.border = Border(

            left=thin_s,

            right=ch.border.right if ch.border else None,

            top=ch.border.top if ch.border else thin_s,

            bottom=ch.border.bottom if ch.border else thin_s,

        )

    if report_type == "SPI" and ws.title == "0001-1":

        # Remove horizontal divider in Top-Right Box between row 3 (label) and row 4 (rep_no)

        thin_s = Side(style="thin", color="000000")

        no_s = Side(style=None)

        for c_idx in range(9, 15):

            c3 = ws.cell(row=3 + offset, column=c_idx)

            c4 = ws.cell(row=4 + offset, column=c_idx)

            l3 = thin_s if c_idx == 9 else no_s

            r3 = thin_s if c_idx == 14 else no_s

            t3 = thin_s

            c3.border = Border(left=l3, right=r3, top=t3, bottom=no_s)

            l4 = thin_s if c_idx == 9 else no_s

            r4 = thin_s if c_idx == 14 else no_s

            b4 = thin_s

            c4.border = Border(left=l4, right=r4, top=no_s, bottom=b4)

    # 6. Drawings Bar

    draw_specs = {

        "FUR": {"type": "combined", "draw": (8, 1), "wmp": (8, 8)},

        "VSR": {"type": "combined", "draw": (8, 1), "wmp": (8, 8)},

        "MTR": {"type": "combined", "draw": (8, 1), "wmp": (8, 7)},

        "DIC": {"type": "separate", "draw_lbl": (9, 1), "draw_val": (9, 4), "wmp_lbl": (9, 10), "wmp_val": (9, 14)},

        "SPI": {"type": "separate", "draw_lbl": (7 if ws.title != "0001-1" else 8, 1), "draw_val": (7 if ws.title != "0001-1" else 8, 3), "wmp_lbl": (7 if ws.title != "0001-1" else 8, 7), "wmp_val": (7 if ws.title != "0001-1" else 8, 9)},

        "CAI": {"type": "separate", "draw_lbl": (7, 1), "draw_val": (7, 3), "wmp_lbl": (7, 7), "wmp_val": (7, 10)},

        "FTI": {"type": "separate", "draw_lbl": (7, 1), "draw_val": (7, 3), "wmp_lbl": (7, 7), "wmp_val": (7, 10)},

        "HTI": {"type": "separate", "draw_lbl": (8, 1), "draw_val": (8, 3), "wmp_lbl": (8, 8), "wmp_val": (8, 10)},

    }

    if report_type in draw_specs:

        ds = draw_specs[report_type]

        if ds["type"] == "combined":

            lbl_draw_font = f9_bi if report_type == "FUR" else f9_i

            val_draw_font = f10_b if report_type == "FUR" else f9_b

            cd = ws.cell(row=ds["draw"][0] + offset, column=ds["draw"][1])

            if not isinstance(cd, MergedCell):

                c_text = clean_draw if clean_draw else ""

                draw_vn = "Bảng vẽ / " if report_type == "FUR" else "Bản vẽ / "

                cd.value = CellRichText(

                    TextBlock(f9_b, draw_vn),

                    TextBlock(lbl_draw_font, "Drawing No.: "),

                    TextBlock(val_draw_font, c_text),

                )

                cd.alignment = align_left

            

            cw = ws.cell(row=ds["wmp"][0] + offset, column=ds["wmp"][1])

            if not isinstance(cw, MergedCell):

                w_text = clean_wmp if clean_wmp else ""

                cw.value = CellRichText(

                    TextBlock(f9_b, "Bản vẽ mối hàn/ "),

                    TextBlock(lbl_draw_font, "Weld map No.: "),

                    TextBlock(val_draw_font, w_text),

                )

                cw.alignment = align_left

        else:

            cd_lbl = ws.cell(row=ds["draw_lbl"][0] + offset, column=ds["draw_lbl"][1])

            if not isinstance(cd_lbl, MergedCell):

                cd_lbl.value = CellRichText(TextBlock(f9_b, "Bản vẽ / "), TextBlock(f9_i, "Drawing No.:"))

                cd_lbl.alignment = align_left

            cd_val = ws.cell(row=ds["draw_val"][0] + offset, column=ds["draw_val"][1])

            if not isinstance(cd_val, MergedCell):

                if clean_draw:

                    cd_val.value = clean_draw

                cd_val.font = font_bold_9

                cd_val.alignment = align_left

            cw_lbl = ws.cell(row=ds["wmp_lbl"][0] + offset, column=ds["wmp_lbl"][1])

            if not isinstance(cw_lbl, MergedCell):

                cw_lbl.value = CellRichText(TextBlock(f9_b, "Bản vẽ mối hàn/ "), TextBlock(f9_i, "Weld map No.:"))

                cw_lbl.alignment = align_left

            cw_val = ws.cell(row=ds["wmp_val"][0] + offset, column=ds["wmp_val"][1])

            if not isinstance(cw_val, MergedCell):

                if clean_wmp:

                    cw_val.value = clean_wmp

                cw_val.font = font_bold_9

                cw_val.alignment = align_left

    if is_spi_0001_1:

        # 1. Date: remove vertical divider between J2 and K2

        for col_l in ["I", "J"]:

            c = ws[f"{col_l}{2 + offset}"]

            c.border = Border(top=c.border.top if c.border else None, bottom=c.border.bottom if c.border else None, left=c.border.left if c.border else None, right=Side(style=None))

        for col_l in ["K", "L", "M", "N"]:

            c = ws[f"{col_l}{2 + offset}"]

            c.border = Border(top=c.border.top if c.border else None, bottom=c.border.bottom if c.border else None, left=Side(style=None), right=c.border.right if c.border else None)

        # 2. Location (B6:D7) and Item (E6:H7): add vertical divider between D and E

        for r_m in [6, 7]:

            r_idx = r_m + offset

            for col_l in ["B", "C", "D"]:

                c = ws[f"{col_l}{r_idx}"]

                r_s = Side(style="thin") if col_l == "D" else (c.border.right if c.border else None)

                c.border = Border(top=c.border.top if c.border else None, bottom=c.border.bottom if c.border else None, left=c.border.left if c.border else None, right=r_s)

            for col_l in ["E", "F", "G", "H"]:

                c = ws[f"{col_l}{r_idx}"]

                l_s = Side(style="thin") if col_l == "E" else (c.border.left if c.border else None)

                c.border = Border(top=c.border.top if c.border else None, bottom=c.border.bottom if c.border else None, left=l_s, right=c.border.right if c.border else None)

        # 3. Drawing No: remove divider between B8 and C8

        for col_l in ["A", "B"]:

            c = ws[f"{col_l}{8 + offset}"]

            c.border = Border(top=c.border.top if c.border else None, bottom=c.border.bottom if c.border else None, left=c.border.left if c.border else None, right=Side(style=None))

        for col_l in ["C", "D", "E", "F"]:

            c = ws[f"{col_l}{8 + offset}"]

            c.border = Border(top=c.border.top if c.border else None, bottom=c.border.bottom if c.border else None, left=Side(style=None), right=c.border.right if c.border else None)

        # 4. Weld map No: remove divider between H8 and I8

        for col_l in ["G", "H"]:

            c = ws[f"{col_l}{8 + offset}"]

            c.border = Border(top=c.border.top if c.border else None, bottom=c.border.bottom if c.border else None, left=c.border.left if c.border else None, right=Side(style=None))

        for col_l in ["I", "J", "K", "L", "M", "N"]:

            c = ws[f"{col_l}{8 + offset}"]

            c.border = Border(top=c.border.top if c.border else None, bottom=c.border.bottom if c.border else None, left=Side(style=None), right=c.border.right if c.border else None)

    # 7. Table Headers for tabular reports (FUR, VSR, MTR)

    table_headers_spec = {

        "FUR": (9, {

            1: ("STT/\n", "No."), 2: ("Số mối hàn/\n", "Joint No."), 3: ("Loại mối hàn/\n", "Joint Type (1)"),

            4: ("Đường kính/\n", "Outside Diameter (mm)"), 5: ("Độ dày/\n", "Thickness (mm)"),

            6: ("Quy cách vật liệu/\n", "Material Specification"), 7: ("Quy trình hàn/\n", "WPS No."),

            8: ("Phương pháp hàn/\n", "Welding Process"), 9: ("Thợ hàn số 1 / Thợ hàn 2/\n", "Welder No. 1/ Welder No. 2"),

            10: ("Làm sạch bên trong ống/\n", "Pipe internal cleaning"), 11: ("Độ thẳng hàn/\n", "Alignment"),

            12: ("Khe hở, vát mép, mép cùn/\n", "Root gap, beveling angle, root face"),

            13: ("Làm sạch tại mối hàn/\n", "Cleaning at welding joint"),

            14: ("Kết quả (Chấp nhận/ Từ chối)/\n", "Results (Acc/Rej)"), 15: ("Ghi chú/\n", "Remarks"),

        }),

        "VSR": (9, {

            1: ("STT/\n", "No."), 2: ("Số mối hàn/\n", "Joint No."), 3: ("Loại mối hàn/\n", "Joint Type (1)"),

            4: ("Đường kính/\n", "Outside Diameter (mm)"), 5: ("Độ dày/\n", "Thickness (mm)"),

            6: ("Quy cách vật liệu/\n", "Material Specification"), 7: ("Quy trình hàn/\n", "WPS No."),

            8: ("Phương pháp hàn/\n", "Welding Process"), 9: ("Thợ hàn số 1 / Thợ hàn 2/\n", "Welder No. 1/ Welder No. 2"),

            10: ("Ngày kiểm tra đấu nối mối hàn/\n", "Fitup Date"), 12: ("Ngày kiểm tra ngoại dạng/\n", "Visual date"),

            14: ("Kết quả (Chấp nhận/ Từ chối)/\n", "Results (Acc/Rej)"), 15: ("Ghi chú/\n", "Remarks"),

        }),

        "MTR": (9, {

            1: ("STT/\n", "No."), 2: ("Số mối hàn/\n", "Joint No."), 3: ("Mô tả/\n", "Description"),

            4: ("Qui cách vật liệu/\n", "Material Specification"), 5: ("Ống số/\n", "Pipe No."),

            7: ("Số nhận dạng/\n", "Serial No."), 9: ("Báo cáo nhận vật tư/\n", "MRIR Report No."),

            12: ("Ghi chú/\n", "Remarks"),

        }),

    }

    if report_type == "FUR":

        fur_table_headers = {

            1: (f7_b, "STT/\n", f7_i, "No."),

            2: (f7_b, "Số mối hàn/\n", f7_i, "Joint No."),

            3: (f7_b, "Loại mối hàn /\n", f7_i, "Joint Type (1)"),

            4: (f7_b, "Đường kính/\n", f7_i, "Outside Diameter (mm)"),

            5: (f7_b, "Độ dày / ", f7_r, "Thickness (mm)"),

            6: (f7_b, "Quy cách vật liệu / ", f7_i, "Material Specification"),

            7: (f7_b, "Quy trình hàn / ", f7_i, "WPS No."),

            8: (f7_b, "Phương pháp hàn / ", f7_i, "Welding Process"),

            9: (f7_r, "Thợ hàn số 1 / Thợ hàn 2 ", f7_i, "Welder No. 1/ Welder No. 2"),

            10: (f7_b, "Làm sạch bên trong ống / ", f7_i, "Pipe internal cleaning"),

            11: (f7_b, "Độ thẳng hàn\n/", f7_bi, "Alignment"),

            12: (f7_b, "Khe hở, vát mép, mép cùn/ ", f7_i, "Root gap, beveling angle, root face"),

            13: (f7_b, "Làm sạch tại mối hàn\n/", f7_i, "Cleaning at welding joint"),

            14: (f7_b, "Kết quả / (Chấp nhận/ Từ chối) ", f7_i, "Results (Acc/Rej)"),

            15: (f7_b, "Ghi chú/\n", f7_i, "Remarks"),

        }

        for col_idx, (f_vn, vn_text, f_en, en_text) in fur_table_headers.items():

            cell = ws.cell(row=9 + offset, column=col_idx)

            if not isinstance(cell, MergedCell):

                cell.value = CellRichText(

                    TextBlock(f_vn, vn_text),

                    TextBlock(f_en, en_text),

                )

                cell.alignment = align_center

    

        # Row 17: Abbreviations & conclusion

        abbr_cell = ws.cell(row=17 + offset, column=1)

        if not isinstance(abbr_cell, MergedCell):

            abbr_cell.value = CellRichText(

                TextBlock(f8_r, "Viết tắt / "),

                TextBlock(f8_i, "Abbreviation"),

                TextBlock(f8_bi, ": "),

                TextBlock(f8_b, "(1) BW"),

                TextBlock(f8_r, ": Hàn giáp mối / "),

                TextBlock(f8_i, "Butt Weld"),

                TextBlock(f8_r, ", "),

                TextBlock(f8_b, "FW"),

                TextBlock(f8_r, ": Hàn góc / "),

                TextBlock(f8_i, "Filet weld"),

                TextBlock(f8_b, ", SV: "),

                TextBlock(f8_r, "vát góc một bên chữ V /"),

                TextBlock(f8_i, "Single V"),

                TextBlock(f8_r, ", "),

                TextBlock(f8_b, "DV"),

                TextBlock(f8_r, ": vát mép hai bên chữ V/"),

                TextBlock(f8_i, "Double V\n"),

                TextBlock(f8_r, 'Kết luận/Conclude:  ☐ ACC - Chấp nhận nghiệm thu, đồng ý cho triển khai công việc tiếp theo/               ☐ REJ\n                                           Acceptance of inspection and approval to proceed with the next phase"\t\t\t\t\t\t\t\t\t\t\t\t\t'),

            )

            abbr_cell.alignment = align_left

    

        # Row 18: Signatures table headers

        f8_i = InlineFont(rFont="Times New Roman", sz=8.0, b=False, i=True)
        if report_type in {"VSR", "FUR", "MTR", "ICR"}:
            ibst_hdr = (f9_b, "GIÁM SÁT THI CÔNG IBST/S\n", f8_i, "IBST/S CONSTRUCTION SUPERVISOR")
        else:
            ibst_hdr = (f9_b, "GIÁM SÁT THI CÔNG IBST/S/\n", f8_r, "IBST/S CONSTRUCTION\nSUPERVISOR")
        sig_hdrs = {
            1: (f9_r, "Kiểm tra/\n", f9_r, "Inspection"),
            3: (f9_b, "TỔNG THẦU EPC - PTSCL18/\n", f9_r, "EPC CONTRACTOR – PTSCL18"),
            6: ibst_hdr,
            9: (f9_b, "CHỦ ĐẦU TƯ-SWPOC/\n", f9_r, "COMPANY-SWPOC"),
            13: (f9_b, "TƯ VẤN ĐĂNG KIỂM/\n", f9_r, "CA"),
        }
        for col_idx, (f_vn, vn_text, f_en, en_text) in sig_hdrs.items():
            cell = ws.cell(row=18 + offset, column=col_idx)
            if not isinstance(cell, MergedCell):
                cell.value = CellRichText(TextBlock(f_vn, vn_text), TextBlock(f_en, en_text))
                cell.alignment = align_center

        # Row 19: Positions
        pos_hdrs = {
            1: (f9_r, "Chức danh/\n", f9_r, "Position"),
            3: (f8_b, "Kỹ thuật thi công/\n", f8_i, "Construction Technician"),
            6: (f8_b, "Giám sát thi công/\n", f8_i, "Construction Supervisor"),
            9: (f8_b, "Đại diện Chủ đầu tư/\n", f8_i, "Company Representative"),
        }
        for col_idx, (f_vn, vn_text, f_en, en_text) in pos_hdrs.items():
            cell = ws.cell(row=19 + offset, column=col_idx)
            if not isinstance(cell, MergedCell):
                cell.value = CellRichText(TextBlock(f_vn, vn_text), TextBlock(f_en, en_text))
                cell.alignment = align_center
        c_ca_pos = ws.cell(row=19 + offset, column=13)
        if not isinstance(c_ca_pos, MergedCell):
            c_ca_pos.value = None

        # Row 20: Chữ ký / Signature
        sig_lbl = ws.cell(row=20 + offset, column=1)
        if not isinstance(sig_lbl, MergedCell):
            sig_lbl.value = CellRichText(TextBlock(f9_r, "Chữ ký/\n"), TextBlock(f9_r, "Signature"))
            sig_lbl.alignment = align_center
        for c_idx in [3, 6, 9, 13]:
            c_clr = ws.cell(row=20 + offset, column=c_idx)
            if not isinstance(c_clr, MergedCell):
                c_clr.value = None

        # Row 21: Họ và tên / Full name
        fn_lbl = ws.cell(row=21 + offset, column=1)
        if not isinstance(fn_lbl, MergedCell):
            fn_lbl.value = CellRichText(TextBlock(f9_r, "Họ và tên/\n"), TextBlock(f9_r, "Full name"))
            fn_lbl.alignment = align_center
        for c_idx in [6, 9, 13]:
            c_clr = ws.cell(row=21 + offset, column=c_idx)
            if not isinstance(c_clr, MergedCell):
                c_clr.value = None
        c_epc_name = ws.cell(row=21 + offset, column=3)
        if not isinstance(c_epc_name, MergedCell):
            if not c_epc_name.value:
                c_epc_name.value = "Nguyễn Đức Sơn"
            c_epc_name.font = Font(name="Times New Roman", size=sz_body, bold=False, italic=False)
            c_epc_name.alignment = align_center

        # Row 22: Note

        note_cell = ws.cell(row=22 + offset, column=1)

        if not isinstance(note_cell, MergedCell):

            note_cell.value = CellRichText(

                TextBlock(f9_b, "Chú thích"),

                TextBlock(f9_r, "/ "),

                TextBlock(f9_i, "Note:"),

                TextBlock(f10_r, "                                    "),

                TextBlock(f10_b, "      Chấp thuận"),

                TextBlock(f10_r, "/ ACC = Aceepted                                                        "),

                TextBlock(f10_b, " Không chấp thuận"),

                TextBlock(f10_r, "/ REJ = Rejected                                                       "),

                TextBlock(f10_b, "Không áp dụng"),

                TextBlock(f10_r, "/ N/A= Not applicable"),

            )

            note_cell.alignment = align_left

    elif report_type in table_headers_spec:

        hdr_row, hdrs = table_headers_spec[report_type]

        for col_idx, (vn_text, en_text) in hdrs.items():

            cell = ws.cell(row=hdr_row + offset, column=col_idx)

            if not isinstance(cell, MergedCell):

                cell.value = CellRichText(

                    TextBlock(f9_b, vn_text),

                    TextBlock(f9_i, en_text),

                )

                cell.alignment = align_center

    # 8. Signatures bar

    if engineer_name:

        for r in range(15, ws.max_row + 1):

            for c in range(1, ws.max_column + 1):

                cell = ws.cell(row=r, column=c)

                if not isinstance(cell, MergedCell) and isinstance(cell.value, str):

                    if any(old in cell.value for old in ["Lê Quốc Đoàn", "Võ Xuân Duy", "Nguyễn Đức Sơn", "Lê Đình Trường"]):

                        cell.value = engineer_name

                        cell.font = font_reg_9

                        cell.alignment = align_center

def fit_spi_joint_columns(sheet, joint_row: int = 11, first_col: int = 9, last_col: int = 14) -> None:

    """Size RP-SPI 0001-1 joint columns: fit columns that contain a joint, shrink empty ones.

    Merged header blocks spanning these columns (Report No., RFI No., weld map, CA signature...)

    impose a minimum total width so their text is never clipped.

    """

    CHAR_W = 1.3  # approx. Excel width units per character at 12pt Times New Roman

    LINE_PT = 15.5  # approx. line height (pt) at 12pt

    EMPTY_W = 3.0

    MIN_JOINT_W = 12.5  # fits "ACC" and times like "10:00AM"

    cols = list(range(first_col, last_col + 1))

    widths: dict[int, float] = {}

    has_joint: dict[int, bool] = {}

    for c in cols:

        val = sheet.cell(row=joint_row, column=c).value

        text = str(val).strip() if val is not None else ""

        if text and not text.startswith("="):

            has_joint[c] = True

            widths[c] = max(MIN_JOINT_W, round(len(text) * 1.35 + 3.0, 1))

        else:

            has_joint[c] = False

            widths[c] = EMPTY_W

    # Minimum widths required by merged blocks fully inside the joint columns

    for rng in sheet.merged_cells.ranges:

        if rng.min_col < first_col or rng.max_col > last_col:

            continue

        val = sheet.cell(row=rng.min_row, column=rng.min_col).value

        if val is None:

            continue

        text = str(val)

        if not text.strip() or text.startswith("="):

            continue

        height = 0.0

        for r in range(rng.min_row, rng.max_row + 1):

            height += sheet.row_dimensions[r].height or 15.0

        pieces = [s.strip() for s in text.split("\n") if s.strip()]

        spare_lines = max(1, int(height // LINE_PT) - len(pieces) + 1)

        longest = max((len(s) for s in pieces), default=0)

        need = longest * CHAR_W / spare_lines + 2.0

        span = list(range(rng.min_col, rng.max_col + 1))

        deficit = need - sum(widths[c] for c in span)

        if deficit <= 0:

            continue

        grow = [c for c in span if not has_joint[c]] or span

        for c in grow:

            widths[c] = round(widths[c] + deficit / len(grow), 2)

    for c in cols:

        sheet.column_dimensions[get_column_letter(c)].width = widths[c]

def restore_template_formatting(

    workbook, template_path: Path, sheet_names: list[str], page_len: int = 0, num_pages: int = 1, paged_sheet: str = "", report_type: str = ""

) -> None:

    """Restore font, column widths, row heights, and page setup from the RP template.

    Values are kept; rich text created by the generator on cells that are plain text in

    the template is flattened so the cell shows exactly the template's formatting.

    """

    try:

        tpl_wb = load_workbook(template_path, data_only=False, rich_text=True)

    except Exception:

        return

    try:

        for name in sheet_names:

            if name not in tpl_wb.sheetnames or name not in workbook.sheetnames:

                continue

            ws_t = tpl_wb[name]

            ws_o = workbook[name]

            # Page Setup & Print Options

            ws_o.page_setup.orientation = ws_t.page_setup.orientation

            ws_o.page_setup.paperSize = ws_t.page_setup.paperSize

            ws_o.page_setup.scale = ws_t.page_setup.scale

            ws_o.page_setup.fitToWidth = ws_t.page_setup.fitToWidth

            ws_o.page_setup.fitToHeight = ws_t.page_setup.fitToHeight

            if ws_t.sheet_properties and ws_t.sheet_properties.pageSetUpPr:
                if not ws_o.sheet_properties:
                    ws_o.sheet_properties = copy.copy(ws_t.sheet_properties)
                elif ws_o.sheet_properties.pageSetUpPr is not None:
                    ws_o.sheet_properties.pageSetUpPr.fitToPage = ws_t.sheet_properties.pageSetUpPr.fitToPage
                else:
                    ws_o.sheet_properties.pageSetUpPr = copy.copy(ws_t.sheet_properties.pageSetUpPr)
            if ws_t.print_options:
                ws_o.print_options.horizontalCentered = ws_t.print_options.horizontalCentered
                ws_o.print_options.verticalCentered = ws_t.print_options.verticalCentered
            if ws_t.print_area:
                if name == paged_sheet and num_pages > 1:
                    parts = ws_t.print_area.split(":")
                    if len(parts) == 2:
                        last_part = re.sub(r'\d+$', str(page_len * num_pages), parts[1])
                        ws_o.print_area = f"{parts[0]}:{last_part}"
                    else:
                        ws_o.print_area = ws_t.print_area
                else:
                    ws_o.print_area = ws_t.print_area
            if name == paged_sheet and num_pages > 1:
                from openpyxl.worksheet.pagebreak import Break
                existing_breaks = {b.id for b in getattr(ws_o.row_breaks, "brk", []) if hasattr(b, "id")}
                for p_brk in range(1, num_pages):
                    brk_row = p_brk * page_len
                    if brk_row not in existing_breaks:
                        ws_o.row_breaks.append(Break(id=brk_row))
            if name == "RP-FTI" or getattr(ws_o, "title", "") == "RP-FTI" or (paged_sheet == "RP-FTI" and name == paged_sheet):
                ws_o.page_setup.fitToWidth = 1
                ws_o.page_setup.fitToHeight = 0
                if not ws_o.sheet_properties:
                    from openpyxl.worksheet.properties import WorksheetProperties
                    ws_o.sheet_properties = WorksheetProperties()
                if ws_o.sheet_properties.pageSetUpPr is None:
                    from openpyxl.worksheet.properties import PageSetUpPr
                    ws_o.sheet_properties.pageSetUpPr = PageSetUpPr()
                ws_o.sheet_properties.pageSetUpPr.fitToPage = True
            elif name == "0001-1" and paged_sheet == "RP-FTI":
                ws_o.page_setup.fitToWidth = 1
                ws_o.page_setup.fitToHeight = 0
                if not ws_o.sheet_properties:
                    from openpyxl.worksheet.properties import WorksheetProperties
                    ws_o.sheet_properties = WorksheetProperties()
                if ws_o.sheet_properties.pageSetUpPr is None:
                    from openpyxl.worksheet.properties import PageSetUpPr
                    ws_o.sheet_properties.pageSetUpPr = PageSetUpPr()
                ws_o.sheet_properties.pageSetUpPr.fitToPage = True
            elif (name == "0001" or getattr(ws_o, "title", "") == "0001") and paged_sheet in {"0001", "RP-MTR"}:
                ws_o.page_setup.fitToWidth = 1
                ws_o.page_setup.fitToHeight = 0
                if not ws_o.sheet_properties:
                    from openpyxl.worksheet.properties import WorksheetProperties
                    ws_o.sheet_properties = WorksheetProperties()
                if ws_o.sheet_properties.pageSetUpPr is None:
                    from openpyxl.worksheet.properties import PageSetUpPr
                    ws_o.sheet_properties.pageSetUpPr = PageSetUpPr()
                ws_o.sheet_properties.pageSetUpPr.fitToPage = True
            if ws_t.sheet_view and ws_t.sheet_view.view:
                ws_o.sheet_view.view = ws_t.sheet_view.view
            # Margins

            ws_o.page_margins.left = ws_t.page_margins.left

            ws_o.page_margins.right = ws_t.page_margins.right

            ws_o.page_margins.top = ws_t.page_margins.top

            ws_o.page_margins.bottom = ws_t.page_margins.bottom

            ws_o.page_margins.header = ws_t.page_margins.header

            ws_o.page_margins.footer = ws_t.page_margins.footer

            # Column widths

            for col_letter, col_dim in ws_t.column_dimensions.items():

                if col_dim.width is not None:

                    ws_o.column_dimensions[col_letter].width = col_dim.width

            rows_t = ws_t.max_row

            cols_t = ws_t.max_column

            offsets = [0]

            if name == paged_sheet and num_pages > 1 and page_len > 0:

                rows_t = min(rows_t, page_len)

                offsets = [(p - 1) * page_len for p in range(1, num_pages + 1)]

            for off in offsets:

                for r in range(1, rows_t + 1):

                    if r in ws_t.row_dimensions and ws_t.row_dimensions[r].height is not None:

                        ws_o.row_dimensions[r + off].height = ws_t.row_dimensions[r].height

                    for c in range(1, cols_t + 1):

                        ct = ws_t.cell(row=r, column=c)

                        co = ws_o.cell(row=r + off, column=c)

                        if isinstance(co, MergedCell) or isinstance(ct, MergedCell):

                            continue

                        if isinstance(co.value, CellRichText):
                            if ct.has_style and ct.font and ct.font.size:
                                for part in co.value:
                                    if isinstance(part, TextBlock) and getattr(part, "font", None):
                                        if getattr(part.font, "sz", 0) and part.font.sz < ct.font.size:
                                            part.font.sz = ct.font.size

                            continue

                        if ct.has_style:

                            co.font = copy.copy(ct.font)

                            if ct.alignment:

                                co.alignment = copy.copy(ct.alignment)

                            if ct.border:

                                co.border = copy.copy(ct.border)

    finally:

        tpl_wb.close()

def _strip_drawing_label(text: str) -> str:

    text = str(text or "").strip()

    if ":" in text:

        text = text.split(":", 1)[1].strip()

    return text

def apply_template_rich_text(

    ws,

    template_path: Path,

    page_len: int,

    num_pages: int,

    station_code: str,

    drawing_no: str,

    weld_map_no: str,

) -> None:

    """Copy the rich-text title/header/footer cells of the RP template 1:1 (fonts, sizes, bold/italic runs).

    Only dynamic pieces are substituted: station code (C7), drawing / weld map numbers (A8, H8)

    and the page number (M6).

    """

    try:

        tpl_wb = load_workbook(template_path, data_only=False, rich_text=True)

    except Exception:

        return

    try:

        if ws.title not in tpl_wb.sheetnames:

            return

        ws_t = tpl_wb[ws.title]

        clean_draw = _strip_drawing_label(drawing_no)

        clean_wmp = _strip_drawing_label(weld_map_no)

        rows_t = min(ws_t.max_row, page_len) if page_len else ws_t.max_row

        for p in range(1, max(1, num_pages) + 1):

            off = (p - 1) * page_len

            for r in range(1, rows_t + 1):

                for c in range(1, min(ws_t.max_column, 30) + 1):

                    ct = ws_t.cell(row=r, column=c)

                    if isinstance(ct, MergedCell) or not isinstance(ct.value, CellRichText):

                        continue

                    co = ws.cell(row=r + off, column=c)

                    if isinstance(co, MergedCell):

                        continue

                    runs = []

                    for part in ct.value:

                        if isinstance(part, TextBlock):

                            runs.append(TextBlock(copy.copy(part.font), part.text))

                        else:

                            runs.append(str(part))

                    def text_of(i: int) -> str:

                        return runs[i].text if isinstance(runs[i], TextBlock) else runs[i]

                    def set_text(i: int, new_text: str) -> None:

                        if isinstance(runs[i], TextBlock):

                            runs[i] = TextBlock(runs[i].font, new_text)

                        else:

                            runs[i] = new_text

                    full = "".join(text_of(i) for i in range(len(runs)))

                    if "Drawing No" in full and clean_draw and any(token in full for token in ["DAG", "DWG", "SHT", "Sheet", "REV", "PLI"]):

                        set_text(len(runs) - 1, clean_draw)

                    elif "Weld map No" in full and clean_wmp and any(token in full for token in ["WMP", "SHT", "Sheet", "REV", "PLI"]):

                        set_text(len(runs) - 1, clean_wmp)

                    elif ("Trang" in full or "Page" in full) and re.search(r"\d+\s*$", full):

                        last = text_of(len(runs) - 1)

                        set_text(len(runs) - 1, re.sub(r"\d+(\s*)$", f"{p}\\1", last))

                    elif "Trạm thi công" in full or "Construction station" in full:

                        for i in range(len(runs)):

                            set_text(i, re.sub(r"TTC15A?", station_code, text_of(i)))

                    co.value = CellRichText(*runs)

                    if ct.has_style:

                        co.font = copy.copy(ct.font)

                        co.alignment = copy.copy(ct.alignment)

    finally:

        tpl_wb.close()

def standardize_all_rp_titles_before_save(workbook, report_type: str, page_len: int = 24, num_pages: int = 1, paged_sheet_title: str = "") -> None:
    """Enforce exact title formatting on all RP worksheets right before saving:
    if report_type == "DIC":
        return
    Line 1 (Vietnamese): Bold, upright, ends with '/\n'
    Line 2 (English): Regular, upright (NOT italic)
    Alignment: center / center with wrap_text=True.
    """
    canonical_titles = {
        "ICR": ("BÁO CÁO KIỂM TRA LÀM SẠCH BÊN TRONG ỐNG/\n", "INTERNAL PIPE CLEANING INSPECTION REPORT"),
        "FUR": ("BÁO CÁO KIỂM TRA ĐẤU NỐI MỐI HÀN/\n", "FIT UP INSPECTION REPORT"),
        "VSR": ("BÁO CÁO KIỂM TRA NGOẠI DẠNG MỐI HÀN/\n", "VISUAL INSPECTION REPORT"),
        "DIC": ("BÁO CÁO KIỂM TRA HIỆU CHUẨN THIẾT BỊ HÀNG NGÀY/\n", "DAILY CALIBRATION/ INSPECTION FOR APPARATUS"),
        "CAI": ("BÁO CÁO KIỂM TRA THI CÔNG SƠN PHỦ/\n", "COATING APPLICATION INSPECTION REPORT"),
        "HTI": ("BÁO CÁO KIỂM TRA HDPUF/\n", "HDPUF TEST INSPECTION REPORT"),
        "MTR": ("BÁO CÁO TRUY SUẤT VẬT TƯ/\n", "MATERIAL TRACEABILITY REPORT"),
        "PPP": ("BÁO CÁO KIỂM TRA KÉO ĐẨY ỐNG XUỐNG HÀO NƯỚC/\n", "INSPECTION REPORT FOR PULLING PUSHING PIPE INTO TRENCH"),
        "MRIR": ("BÁO CÁO KIỂM TRA VÀ NHẬN VẬT TƯ/\n", "MATERIAL RECEIVING INSPECTION REPORT (MRIR)"),
    }

    def get_rep_title(rep_type, sheet_title):
        if rep_type == "SPI":
            if sheet_title == "0001-1":
                return ("DANH MỤC KIỂM TRA TÌNH TRẠNG BỀ MẶT THÉP/\n", "CHECKLIST RECORD OF STEEL SURFACE CONDITION")
            return ("BÁO CÁO KIỂM TRA CHUẨN BỊ BỀ MẶT/\n", "SURFACE PREPARATION INSPECTION REPORT")
        elif rep_type == "FTI":
            if sheet_title == "0001-1":
                return ("BÁO CÁO KIỂM TRA SỬA CHỮA MỐI BỌC/\n", "COATING JOINT REPAIR INSPECTION REPORT")
            return ("BÁO CÁO KIỂM TRA HIỆN TRƯỜNG/\n", "FIELD JOINT COATING INSPECTION REPORT")
        return canonical_titles.get(rep_type, None)

    
    def standardize_signatures_on_sheet(ws_target, p_len, n_pages, rep_type: str = ""):
        if rep_type == "PPP" or report_type == "PPP":
            return
        sz_sig = (12.0 if getattr(ws_target, "title", "") == "0001-1" else 11.0) if (rep_type == "SPI" or report_type == "SPI") else (11.0 if (rep_type in {"DIC"} or report_type in {"DIC"}) else 10.0)
        for r_probe in range(15, min(ws_target.max_row + 1, 45)):
            c_probe = ws_target.cell(row=r_probe, column=1)
            v_p = str(c_probe.value or "").strip()
            if any(k in v_p for k in ["Kiểm tra", "Chức danh", "Chữ ký"]):
                if c_probe.font and c_probe.font.size and c_probe.font.size >= 10:
                    if rep_type != "SPI" and report_type != "SPI":
                        sz_sig = c_probe.font.size
                    break
        f_r = InlineFont(rFont="Times New Roman", sz=sz_sig, b=False, i=False)
        f_b = InlineFont(rFont="Times New Roman", sz=sz_sig, b=True, i=False)
        f_i = InlineFont(rFont="Times New Roman", sz=sz_sig, b=False, i=True)
        cell_font_r = Font(name="Times New Roman", size=sz_sig, bold=False, italic=False)
        cell_font_b = Font(name="Times New Roman", size=sz_sig, bold=True, italic=False)
        align_c = Alignment(horizontal="center", vertical="center", wrap_text=True)

        v_kiemtra = CellRichText(TextBlock(f_r, "Kiểm tra/\n"), TextBlock(f_r, "Inspection"))
        v_epc = CellRichText(TextBlock(f_b, "TỔNG THẦU EPC - PTSCL18/\n"), TextBlock(f_r, "EPC CONTRACTOR – PTSCL18"))
        if report_type in {"VSR", "FUR", "MTR", "ICR"} or rep_type in {"VSR", "FUR", "MTR", "ICR"}:
            v_ibst = CellRichText(TextBlock(f_b, "GIÁM SÁT THI CÔNG IBST/S\n"), TextBlock(f_i, "IBST/S CONSTRUCTION SUPERVISOR"))
        else:
            v_ibst = CellRichText(TextBlock(f_b, "GIÁM SÁT THI CÔNG IBST/S/\n"), TextBlock(f_r, "IBST/S CONSTRUCTION\nSUPERVISOR"))
        v_swpoc = CellRichText(TextBlock(f_b, "CHỦ ĐẦU TƯ-SWPOC/\n"), TextBlock(f_r, "COMPANY-SWPOC"))
        v_ca = CellRichText(TextBlock(f_b, "TƯ VẤN ĐĂNG KIỂM/\n"), TextBlock(f_r, "CA"))

        v_chucdanh = CellRichText(TextBlock(f_r, "Chức danh/\n"), TextBlock(f_r, "Position"))
        v_pos_epc = CellRichText(TextBlock(f_b, "Kỹ thuật thi công/\n"), TextBlock(f_i, "Construction Technician"))
        v_pos_ibst = CellRichText(TextBlock(f_b, "Giám sát thi công/\n"), TextBlock(f_i, "Construction Supervisor"))
        v_pos_swpoc = CellRichText(TextBlock(f_b, "Đại diện Chủ đầu tư/\n"), TextBlock(f_i, "Company Representative"))

        v_chuky = CellRichText(TextBlock(f_r, "Chữ ký/\n"), TextBlock(f_r, "Signature"))
        v_hovaten = CellRichText(TextBlock(f_r, "Họ và tên/\n"), TextBlock(f_r, "Full name"))

        for p_idx in range(1, max(1, n_pages) + 1):
            offset_p = (p_idx - 1) * p_len
            rows_by_label = {}
            min_scan_r = (22 if (report_type in {"FTI", "SPI", "CAI", "HTI"} or rep_type in {"FTI", "SPI", "CAI", "HTI"}) else 15) + offset_p
            for r_scan in range(min_scan_r, min(ws_target.max_row + 1, 45 + offset_p)):
                c_scan = ws_target.cell(row=r_scan, column=1)
                if isinstance(c_scan, MergedCell):
                    continue
                v_scan = str(c_scan.value or "").strip()
                if not v_scan:
                    continue
                if any(k in v_scan for k in ["Kiểm tra", "Inspection"]) and not any(k in v_scan for k in ["muối", "trực quan", "kích thước", "BÁO CÁO", "khuyết tật", "Holiday", "lớp bọc", "sơn", "bề mặt", "HDPUF", "PE", "đấu nối"]):
                    rows_by_label["kiemtra"] = r_scan
                    c_scan.value = v_kiemtra
                    c_scan.font = cell_font_r
                    c_scan.alignment = align_c
                elif any(k in v_scan for k in ["Chức danh", "Position"]):
                    rows_by_label["chucdanh"] = r_scan
                    c_scan.value = v_chucdanh
                    c_scan.font = cell_font_r
                    c_scan.alignment = align_c
                elif any(k in v_scan for k in ["Chữ ký", "Signature"]):
                    rows_by_label["chuky"] = r_scan
                    c_scan.value = v_chuky
                    c_scan.font = cell_font_r
                    c_scan.alignment = align_c
                elif any(k in v_scan for k in ["Họ và tên", "Full name"]):
                    rows_by_label["hovaten"] = r_scan
                    c_scan.value = v_hovaten
                    c_scan.font = cell_font_r
                    c_scan.alignment = align_c

            cols = {}
            if "kiemtra" in rows_by_label:
                r_kt = rows_by_label["kiemtra"]
                for c_scan in range(2, min(ws_target.max_column + 1, 20)):
                    cell_pt = ws_target.cell(row=r_kt, column=c_scan)
                    if isinstance(cell_pt, MergedCell):
                        continue
                    v_pt = str(cell_pt.value or "").strip()
                    if "TỔNG THẦU" in v_pt or "PTSCL" in v_pt:
                        cols["epc"] = c_scan
                        cell_pt.value = v_epc
                        cell_pt.font = cell_font_b
                        cell_pt.alignment = align_c
                    elif "IBST" in v_pt:
                        cols["ibst"] = c_scan
                        cell_pt.value = v_ibst
                        cell_pt.font = cell_font_b
                        cell_pt.alignment = align_c
                    elif "SWPOC" in v_pt or "CHỦ ĐẦU TƯ" in v_pt:
                        cols["swpoc"] = c_scan
                        cell_pt.value = v_swpoc
                        cell_pt.font = cell_font_b
                        cell_pt.alignment = align_c
                    elif "TƯ VẤN ĐĂNG KIỂM" in v_pt or "CA" in v_pt:
                        cols["ca"] = c_scan
                        cell_pt.value = v_ca
                        cell_pt.font = cell_font_b
                        cell_pt.alignment = align_c

            if "chucdanh" in rows_by_label and ("kiemtra" not in rows_by_label or not cols):
                r_cd_scan = rows_by_label["chucdanh"]
                for c_scan in range(2, min(ws_target.max_column + 1, 20)):
                    cell_pt = ws_target.cell(row=r_cd_scan, column=c_scan)
                    if isinstance(cell_pt, MergedCell):
                        continue
                    v_pt = str(cell_pt.value or "").strip()
                    if "thi công" in v_pt or "Technician" in v_pt:
                        cols["epc"] = c_scan
                        cell_pt.value = v_pos_epc
                        cell_pt.font = cell_font_r
                        cell_pt.alignment = align_c
                    elif "Giám sát" in v_pt or "Supervisor" in v_pt:
                        cols["ibst"] = c_scan
                        cell_pt.value = v_pos_ibst
                        cell_pt.font = cell_font_r
                        cell_pt.alignment = align_c
                    elif "Chủ đầu tư" in v_pt or "Representative" in v_pt:
                        cols["swpoc"] = c_scan
                        cell_pt.value = v_pos_swpoc
                        cell_pt.font = cell_font_r
                        cell_pt.alignment = align_c

            if "chucdanh" in rows_by_label:
                r_cd = rows_by_label["chucdanh"]
                if "epc" in cols:
                    c_cell = ws_target.cell(row=r_cd, column=cols["epc"])
                    if not isinstance(c_cell, MergedCell):
                        c_cell.value = v_pos_epc
                        c_cell.font = cell_font_r
                        c_cell.alignment = align_c
                if "ibst" in cols:
                    c_cell = ws_target.cell(row=r_cd, column=cols["ibst"])
                    if not isinstance(c_cell, MergedCell):
                        c_cell.value = v_pos_ibst
                        c_cell.font = cell_font_r
                        c_cell.alignment = align_c
                if "swpoc" in cols:
                    c_cell = ws_target.cell(row=r_cd, column=cols["swpoc"])
                    if not isinstance(c_cell, MergedCell):
                        c_cell.value = v_pos_swpoc
                        c_cell.font = cell_font_r
                        c_cell.alignment = align_c
                if "ca" in cols:
                    c_cell = ws_target.cell(row=r_cd, column=cols["ca"])
                    if not isinstance(c_cell, MergedCell):
                        c_cell.value = None
                        c_cell.font = cell_font_r

            if "chuky" in rows_by_label:
                r_ck = rows_by_label["chuky"]
                for c_idx in cols.values():
                    c_cell = ws_target.cell(row=r_ck, column=c_idx)
                    if not isinstance(c_cell, MergedCell):
                        c_cell.value = None
                        c_cell.font = cell_font_r

            if "hovaten" in rows_by_label:
                r_ten = rows_by_label["hovaten"]
                for k_name, c_idx in cols.items():
                    c_cell = ws_target.cell(row=r_ten, column=c_idx)
                    if not isinstance(c_cell, MergedCell):
                        if k_name == "epc":
                            if not c_cell.value:
                                c_cell.value = "Nguyễn Đức Sơn"
                            c_cell.font = Font(name="Times New Roman", size=sz_sig, bold=False, italic=False)
                            c_cell.alignment = align_c
                        else:
                            c_cell.value = None
                            c_cell.font = cell_font_r

    for s_name in workbook.sheetnames:
        if s_name.startswith("Controlog"):
            continue
        ws = workbook[s_name]
        n_pages = num_pages if (paged_sheet_title and s_name == paged_sheet_title) else 1
        p_len = page_len if page_len else 24

        for p in range(1, max(1, n_pages) + 1):
            off = (p - 1) * p_len
            for r in range(1, min(ws.max_row, 10)):
                for c in range(1, min(ws.max_column, 12)):
                    cell = ws.cell(row=r + off, column=c)
                    if isinstance(cell, MergedCell):
                        continue
                    val = cell.value
                    if not val:
                        continue
                    s = str(val)

                    # Project Title cell
                    if any(k in s for k in ["DỰ ÁN ĐƯỜNG ỐNG", "DỰ ÁN ĐƯỜNG ỐNG DẪN KHÍ", "GAS PIPELINE PROJECT", "ĐƯỜNG ỐNG DẪN KHÍ LÔ B"]) and "BÁO CÁO" not in s:
                        sz = 14.0 if report_type == "SPI" else (12.0 if report_type in {"CAI", "FTI", "HTI", "PPP", "MRIR"} else 10.0)
                        if isinstance(val, CellRichText):
                            for part in val:
                                if isinstance(part, TextBlock) and getattr(part.font, "sz", None):
                                    sz = part.font.sz
                                    break
                        elif cell.font and cell.font.size:
                            sz = cell.font.size

                        f_b = InlineFont(rFont="Times New Roman", sz=sz, b=True, i=False)
                        f_r = InlineFont(rFont="Times New Roman", sz=sz, b=False, i=False)
                        cell.value = CellRichText(
                            TextBlock(f_b, "DỰ ÁN ĐƯỜNG ỐNG DẪN KHÍ LÔ B- Ô MÔN/\n"),
                            TextBlock(f_r, "BLOCK B – O MON GAS PIPELINE PROJECT")
                        )
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                    # Report Title cell
                    elif any(k in s for k in ["BÁO CÁO", "DANH MỤC", "REPORT", "CHECKLIST"]) and not any(k in s for k in ["Biểu mẫu", "Form No", "RFI", "MRIR No"]):
                        sz = 14.0 if report_type == "SPI" else (12.0 if report_type in {"CAI", "FTI", "HTI", "PPP", "MRIR"} else 10.0)
                        if isinstance(val, CellRichText):
                            for part in val:
                                if isinstance(part, TextBlock) and getattr(part.font, "sz", None):
                                    sz = part.font.sz
                                    break
                        elif cell.font and cell.font.size:
                            sz = cell.font.size

                        t_pair = get_rep_title(report_type, ws.title)
                        if t_pair:
                            vn_t, en_t = t_pair
                        else:
                            if "/" in s:
                                parts = s.split("/", 1)
                                vn_t = parts[0].strip() + "/\n"
                                en_t = parts[1].strip()
                            elif "\n" in s:
                                parts = s.split("\n", 1)
                                vn_t = parts[0].strip() + "/\n"
                                en_t = parts[1].strip()
                            else:
                                vn_t = s.strip() + "/\n"
                                en_t = ""

                        f_b = InlineFont(rFont="Times New Roman", sz=sz, b=True, i=False)
                        f_r = InlineFont(rFont="Times New Roman", sz=sz, b=False, i=False)
                        cell.value = CellRichText(
                            TextBlock(f_b, vn_t),
                            TextBlock(f_r, en_t)
                        )
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                    # Location cell
                    elif any(k in s for k in ["Trạm thi công", "Construction station"]) and ("Địa điểm" not in s and "Location" not in s or "/" in s):
                        st = "TTC15"
                        if "TTC11" in s:
                            st = "TTC11"
                        elif "TTC15A" in s:
                            st = "TTC15A"
                        elif "TTC15" in s:
                            st = "TTC15"
                        sz_loc = (12 if getattr(ws, "title", "") == "0001-1" else 11) if report_type == "SPI" else 9
                        f_r = InlineFont(rFont="Times New Roman", sz=sz_loc, b=False, i=False)
                        f_i = InlineFont(rFont="Times New Roman", sz=sz_loc, b=False, i=True)
                        cell.value = CellRichText(
                            TextBlock(f_r, f"Trạm thi công {st} - An Giang/\n"),
                            TextBlock(f_i, f"Construction station {st} - An Giang")
                        )
                        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                    # Item cell
                    elif any(k in s for k in ["Thi công tuyến ống bờ", "Onshore pipeline"]) and ("Hạng mục" not in s and "Item" not in s or "/" in s):
                        sz_item = (12 if getattr(ws, "title", "") == "0001-1" else 11) if report_type == "SPI" else 9
                        f_r = InlineFont(rFont="Times New Roman", sz=sz_item, b=False, i=False)
                        f_i = InlineFont(rFont="Times New Roman", sz=sz_item, b=False, i=True)
                        cell.value = CellRichText(
                            TextBlock(f_r, "Thi công tuyến ống bờ/\n"),
                            TextBlock(f_i, "Onshore pipeline")
                        )
                        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        standardize_signatures_on_sheet(ws, p_len, n_pages, report_type)

        if report_type == "CAI":
            for p in range(1, max(1, n_pages) + 1):
                off = (p - 1) * p_len
                i_j_range = f"I{2 + off}:J{2 + off}"
                k_n_range = f"K{2 + off}:N{2 + off}"
                existing_merges = {str(m) for m in ws.merged_cells.ranges}
                if i_j_range not in existing_merges:
                    ws.merge_cells(start_row=2 + off, start_column=9, end_row=2 + off, end_column=10)
                if k_n_range not in existing_merges:
                    ws.merge_cells(start_row=2 + off, start_column=11, end_row=2 + off, end_column=14)
                c_lbl = ws.cell(row=2 + off, column=9)
                c_lbl.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
                c_val = ws.cell(row=2 + off, column=11)
                c_val.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
                for r in [12, 13, 14, 15, 16, 17, 18, 19]:
                    cell = ws.cell(row=r + off, column=1)
                    if not isinstance(cell, MergedCell):
                        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                        cell.font = Font(name="Times New Roman", size=10.0, bold=False, italic=False, color="000000")

        if report_type == "DIC":
            for p in range(1, max(1, n_pages) + 1):
                off = (p - 1) * p_len
                for r_offset in [7, 8]:
                    r_idx = r_offset + off
                    for col_idx in range(15, 19):
                        ws.cell(row=r_idx, column=col_idx).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        if report_type == "FTI":
            for p in range(1, max(1, n_pages) + 1):
                off = (p - 1) * p_len
                # J4: Report No chữ thường, không in đậm (bold=False), size 9pt
                c_rep = ws.cell(row=4 + off, column=10)
                if not isinstance(c_rep, MergedCell):
                    c_rep.font = Font(name="Times New Roman", size=9.0, bold=False, italic=False, color="000000")

                # Xóa gạch: Remove horizontal divider line between row 3 and row 4 for cols J..O (10..15)
                for c_sep in range(10, 16):
                    c_r3 = ws.cell(row=3 + off, column=c_sep)
                    c_r4 = ws.cell(row=4 + off, column=c_sep)
                    if c_r3.border:
                        c_r3.border = Border(left=c_r3.border.left, right=c_r3.border.right, top=c_r3.border.top, bottom=None)
                    if c_r4.border:
                        c_r4.border = Border(left=c_r4.border.left, right=c_r4.border.right, top=None, bottom=c_r4.border.bottom)

                # A6, D6: bold 10pt
                for col_idx in [1, 4]:
                    c_hdr = ws.cell(row=6 + off, column=col_idx)
                    if not isinstance(c_hdr, MergedCell):
                        c_hdr.font = Font(name="Times New Roman", size=10.0, bold=True, italic=False, color="000000")
                # J6 (RFI No): bold 9pt
                c_rfi = ws.cell(row=6 + off, column=10)
                if not isinstance(c_rfi, MergedCell):
                    c_rfi.font = Font(name="Times New Roman", size=9.0, bold=True, italic=False, color="000000")
                # C7 (Drawing No) & J7 (Weld Map No): bold 9pt
                for col_idx in [3, 10]:
                    c_dwg = ws.cell(row=7 + off, column=col_idx)
                    if not isinstance(c_dwg, MergedCell):
                        c_dwg.font = Font(name="Times New Roman", size=9.0, bold=True, italic=False, color="000000")

                # Row 11: A11 chữ thường, không in đậm, không nghiêng (bold=False, italic=False), 10pt
                c_11 = ws.cell(row=11 + off, column=1)
                if not isinstance(c_11, MergedCell):
                    c_11.value = "Kiểm tra kích thước lớp bọc HTLP80-HP/\nCheck dimension of HTLP80-HP"
                    c_11.font = Font(name="Times New Roman", size=10.0, bold=False, italic=False, color="000000")
                    c_11.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                # Row 15: strictly preserve "Kiểm tra khuyết tật lớp bọc/Holiday Detection test", regular 10pt
                c_15 = ws.cell(row=15 + off, column=1)
                if not isinstance(c_15, MergedCell):
                    c_15.value = "Kiểm tra khuyết tật lớp bọc/Holiday Detection test"
                    c_15.font = Font(name="Times New Roman", size=10.0, bold=False, italic=False, color="000000")
                    c_15.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                # Rows 12..16: ensure left aligned, 10pt regular for column 1 (except row 11)
                for r_item in [12, 13, 14, 16]:
                    c_item = ws.cell(row=r_item + off, column=1)
                    if not isinstance(c_item, MergedCell):
                        c_item.font = Font(name="Times New Roman", size=10.0, bold=False, italic=False, color="000000")
                        c_item.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            if (ws.title == paged_sheet_title or ws.title == "RP-FTI"):
                ws.print_area = f"A1:O{p_len * n_pages}"
                ws.page_setup.orientation = "portrait"
                ws.page_setup.paperSize = 9
                ws.page_setup.scale = 64
                ws.page_setup.fitToWidth = 1
                ws.page_setup.fitToHeight = 0
                if not ws.sheet_properties:
                    from openpyxl.worksheet.properties import WorksheetProperties
                    ws.sheet_properties = WorksheetProperties()
                if ws.sheet_properties.pageSetUpPr is None:
                    from openpyxl.worksheet.properties import PageSetUpPr
                    ws.sheet_properties = PageSetUpPr()
                ws.sheet_properties.pageSetUpPr.fitToPage = True
                ws.page_margins.left = 0.7086614173228347
                ws.page_margins.right = 0.3149606299212598
                ws.page_margins.top = 0.3149606299212598
                ws.page_margins.bottom = 0.3149606299212598
                ws.page_margins.header = 0.0
                ws.page_margins.footer = 0.0
                ws.print_options.horizontalCentered = True
                ws.print_options.verticalCentered = True

                if n_pages > 1:
                    from openpyxl.worksheet.pagebreak import Break
                    existing_breaks = {b.id for b in getattr(ws.row_breaks, "brk", []) if hasattr(b, "id")}
                    for p_brk in range(1, n_pages):
                        brk_row = p_brk * p_len
                        if brk_row not in existing_breaks:
                            ws.row_breaks.append(Break(id=brk_row))
            elif ws.title == "0001-1":
                ws.page_setup.orientation = "portrait"
                ws.page_setup.paperSize = 9
                ws.page_setup.scale = 62
                ws.page_setup.fitToWidth = 1
                ws.page_setup.fitToHeight = 0
                if not ws.sheet_properties:
                    from openpyxl.worksheet.properties import WorksheetProperties
                    ws.sheet_properties = WorksheetProperties()
                if ws.sheet_properties.pageSetUpPr is None:
                    from openpyxl.worksheet.properties import PageSetUpPr
                    ws.sheet_properties = PageSetUpPr()
                ws.sheet_properties.pageSetUpPr.fitToPage = True
                ws.page_margins.left = 0.7086614173228347
                ws.page_margins.right = 0.3149606299212598
                ws.page_margins.top = 0.3149606299212598
                ws.page_margins.bottom = 0.3149606299212598
                ws.page_margins.header = 0.0
                ws.page_margins.footer = 0.0
                ws.print_options.horizontalCentered = True
                ws.print_options.verticalCentered = True

        elif report_type == "MTR" and (ws.title == paged_sheet_title or ws.title == "0001"):
            ws.print_area = f"A1:L{22 * n_pages}"
            ws.page_setup.orientation = "landscape"
            ws.page_setup.paperSize = 9
            ws.page_setup.scale = 87
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            if not ws.sheet_properties:
                from openpyxl.worksheet.properties import WorksheetProperties
                ws.sheet_properties = WorksheetProperties()
            if ws.sheet_properties.pageSetUpPr is None:
                from openpyxl.worksheet.properties import PageSetUpPr
                ws.sheet_properties = PageSetUpPr()
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_margins.left = 0.25
            ws.page_margins.right = 0.25
            ws.page_margins.top = 0.82
            ws.page_margins.bottom = 0.25
            ws.page_margins.header = 0.25
            ws.page_margins.footer = 0.0
            ws.print_options.horizontalCentered = True
            ws.print_options.verticalCentered = False
            if n_pages > 1:
                from openpyxl.worksheet.pagebreak import Break
                existing_breaks = {b.id for b in getattr(ws.row_breaks, "brk", []) if hasattr(b, "id")}
                for p_brk in range(1, n_pages):
                    brk_row = p_brk * 22
                    if brk_row not in existing_breaks:
                        ws.row_breaks.append(Break(id=brk_row))

        elif report_type == "PPP" and (ws.title == paged_sheet_title or ws.title == "0001"):
            ws.print_area = f"'0001'!$A$1:$K${22 * n_pages}"
            ws.page_setup.orientation = "landscape"
            ws.page_setup.paperSize = 9
            ws.page_setup.scale = 77
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            if not ws.sheet_properties:
                from openpyxl.worksheet.properties import WorksheetProperties
                ws.sheet_properties = WorksheetProperties()
            if ws.sheet_properties.pageSetUpPr is None:
                from openpyxl.worksheet.properties import PageSetUpPr
                ws.sheet_properties = PageSetUpPr()
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_margins.left = 0.27
            ws.page_margins.right = 0.16
            ws.page_margins.top = 0.62
            ws.page_margins.bottom = 0.2
            ws.page_margins.header = 0.0
            ws.page_margins.footer = 0.0
            ws.print_options.horizontalCentered = True
            ws.print_options.verticalCentered = False
            if n_pages > 1:
                from openpyxl.worksheet.pagebreak import Break
                existing_breaks = {b.id for b in getattr(ws.row_breaks, "brk", []) if hasattr(b, "id")}
                for p_brk in range(1, n_pages):
                    brk_row = p_brk * 22
                    if brk_row not in existing_breaks:
                        ws.row_breaks.append(Break(id=brk_row))

                # Clone header logos to page 2..n
                if n_pages > 1:
                    base_imgs = [img for img in getattr(ws, "_images", []) if getattr(img.anchor, "_from", None) and img.anchor._from.row < p_len]
                    existing_p2 = [img for img in getattr(ws, "_images", []) if getattr(img.anchor, "_from", None) and img.anchor._from.row >= p_len]
                    if not existing_p2 and base_imgs:
                        import io
                        from openpyxl.drawing.image import Image as OpenPyXLImage
                        from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
                        cloned_imgs = []
                        for p_idx in range(2, n_pages + 1):
                            p_off = (p_idx - 1) * p_len
                            for b_img in base_imgs:
                                if isinstance(b_img.anchor, TwoCellAnchor):
                                    img_bytes = io.BytesIO(b_img.ref.getvalue())
                                    p_img = OpenPyXLImage(img_bytes)
                                    p_img.width = b_img.width
                                    p_img.height = b_img.height
                                    frm = AnchorMarker(
                                        col=b_img.anchor._from.col,
                                        colOff=b_img.anchor._from.colOff,
                                        row=b_img.anchor._from.row + p_off,
                                        rowOff=b_img.anchor._from.rowOff
                                    )
                                    to_m = AnchorMarker(
                                        col=b_img.anchor.to.col,
                                        colOff=b_img.anchor.to.colOff,
                                        row=b_img.anchor.to.row + p_off,
                                        rowOff=b_img.anchor.to.rowOff
                                    )
                                    p_img.anchor = TwoCellAnchor(_from=frm, to=to_m, editAs=b_img.anchor.editAs)
                                    cloned_imgs.append(p_img)
                        ws._images = list(ws._images) + cloned_imgs

        if report_type == "DIC":
            has_dual_logo = any(img.width == 280 for img in getattr(ws, "_images", []))
            if not has_dual_logo:
                base_dir = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else "."
                logo_file = os.path.join(base_dir, "Fome_Web_Data", "ptsc_lilama_logo.png")
                if not os.path.exists(logo_file):
                    logo_file = os.path.join("Fome_Web_Data", "ptsc_lilama_logo.png")
                if os.path.exists(logo_file):
                    from openpyxl.drawing.image import Image as OpenPyXLImage
                    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
                    from openpyxl.drawing.xdr import XDRPositiveSize2D
                    keep_imgs = [img for img in getattr(ws, "_images", []) if img.width != 77]
                    dual_img = OpenPyXLImage(logo_file)
                    dual_img.width = 280
                    dual_img.height = 40
                    marker = AnchorMarker(col=14, colOff=120000, row=0, rowOff=60000)
                    size = XDRPositiveSize2D(cx=280 * 9525, cy=40 * 9525)
                    dual_img.anchor = OneCellAnchor(_from=marker, ext=size)
                    ws._images = keep_imgs + [dual_img]

def generate_rp_workbook(

    rfi_id: str,

    form_rows: dict[str, dict[int, object]],

    external_tables: dict[str, list[tuple[int, tuple[object, ...]]]],

    location: str = DEFAULT_SITE,

    site_cfg: dict[str, str] | None = None,

) -> tuple[Path, str]:

    location = validate_site_code(location, require_existing=True)

    if rfi_id not in form_rows:

        matched_id = None

        for k, row in form_rows.items():

            if str(row.get(13) or "").strip() == rfi_id:

                matched_id = k

                break

        if matched_id:

            rfi_id = matched_id

    report_type = get_rfi_type(rfi_id)

    if not report_type:

        raise UserError(f"Không xác định được loại RP từ số RFI {rfi_id}.")

    template = find_rp_template(report_type, location)

    if template is None:

        raise UserError(f"Chưa có biểu mẫu RP cho loại {report_type}.")

    if rfi_id not in form_rows:

        raise UserError(f"Không tìm thấy số RFI {rfi_id} trong file FORM.")

    if report_type not in RP_SELECTORS and report_type != "MRIR":

        raise UserError(f"Chưa cấu hình ô chọn RFI cho biểu mẫu RP-{report_type}.")

    try:

        workbook = load_workbook(

            template,

            data_only=False,

            keep_vba=template.suffix.lower() == ".xlsm",

            keep_links=True,

        )

    except Exception as exc:

        raise UserError(f"Không đọc được biểu mẫu {template.name}: {exc}") from exc

    helper_values: dict[tuple[str, str], object] = {}

    control_table = []

    pipe_table = []

    if report_type == "ICR" or report_type == "MRIR" or report_type in RP_ASSIGNMENT_COLUMNS:

        ctrl_key = f"Control Log  {location}" if f"Control Log  {location}" in external_tables else "Control Log  TTC15A"

        pipe_key = f"Pipe {location}" if f"Pipe {location}" in external_tables else "Pipe TTC15A"

        control_table = external_tables.get(ctrl_key) or []

        pipe_table = external_tables.get(pipe_key) or []

        helper_values = build_rp_helper_values(

            workbook, report_type, control_table, pipe_table

        )

    # Detect main report sheet

    target_sheet = None

    for s in workbook.worksheets:

        if s.title.casefold().startswith("controlog"):

            continue

        if any(k in s.title.upper() for k in ["RP-", "MRIR", report_type]):

            target_sheet = s

            break

    if not target_sheet:

        target_sheet = workbook.worksheets[0]

    selector = RP_SELECTORS.get(report_type)

    source_row = form_rows[rfi_id]

    if selector:

        helper_values[(target_sheet.title, selector)] = rfi_id

        target_cell = target_sheet[selector]
        if isinstance(target_cell, MergedCell):
            tl_cell = None
            for m in target_sheet.merged_cells.ranges:
                if m.min_row <= target_cell.row <= m.max_row and m.min_col <= target_cell.column <= m.max_col:
                    tl_cell = target_sheet.cell(row=m.min_row, column=m.min_col)
                    break
            if tl_cell is not None:
                tl_cell.value = rfi_id
                helper_values[(target_sheet.title, tl_cell.coordinate)] = rfi_id
        else:
            target_cell.value = rfi_id

    engine = RPFormulaEngine(

        workbook,

        form_rows,

        external_tables,

        helper_values,

        rfi_id,

    )

    station_code = station_code_for_location(location)

    if site_cfg is None:

        site_cfg = get_site_config(location)

    dynamic_drawing_no, dynamic_weld_map_no = resolve_drawings_for_rfi(

        rfi_id,

        report_type,

        external_tables,

        location,

        fallback_drawing=site_cfg["drawing_no"],

        fallback_weld_map=site_cfg["weld_map_no"],

    )

    if report_type in {"ICR", "DIC"}:

        dynamic_drawing_no = re.sub(r"\s+(?:SHT|Sheet).*$", "", dynamic_drawing_no, flags=re.IGNORECASE).strip()

        dynamic_weld_map_no = re.sub(r"\s+(?:SHT|Sheet).*$", "", dynamic_weld_map_no, flags=re.IGNORECASE).strip()

    items = get_rfi_report_items(report_type, rfi_id, control_table, pipe_table)

    if not items and rfi_id in form_rows:

        row_data = form_rows[rfi_id]

        joints_text = row_data.get(29) or row_data.get(12) or ""

        if "joint:" in str(joints_text).lower():

            joints_text = str(joints_text).split("joint:", 1)[1]

        parsed = [j.strip() for j in re.split(r"[;,]+", str(joints_text)) if j.strip() and "KP" in j]

        if parsed:

            items = parsed

    total_items = len(items)

    layout = detect_rp_layout(target_sheet)

    diag_side = Side(style="thin", color="000000")

    num_pages = 1

    page_len = 24

    empty_row_ranges: list[tuple[int, int, int, int]] = []

    engineer_name_rp = (

        "Nguyễn Đức Sơn"

        if ("TTC15" in location.upper() or "T15" in location.upper())

        else ((site_cfg.get("engineer_name") if site_cfg else None) or default_engineer_for_location(location))

    )

    # =========================================================================

    # CASE 1: HORIZONTAL GEOMETRY (FTI, HTI, SPI, CAI)

    # =========================================================================

    if report_type in HORIZONTAL_RP_GEOMETRY:

        capacity = HORIZONTAL_RP_GEOMETRY[report_type].get(target_sheet.title, {}).get("max_joints", 6)

        num_pages = max(1, (total_items + capacity - 1) // capacity) if total_items > 0 else 1

        r_data = form_rows.get(rfi_id, {})

        rep_no = str(r_data.get(13) or "").strip()

        d_val = r_data.get(16) or r_data.get(14) or r_data.get(27)

        d_str = format_rp_date_str(d_val)

        # Update all report sheets (e.g. RP-FTI and 0001-1)
        target_page_len = 31 if (report_type in {"CAI", "FTI", "HTI", "SPI"} or target_sheet.title == f"RP-{report_type}") else 24
        for s_name in [target_sheet.title, "0001-1", "0001"]:

            if s_name not in workbook.sheetnames:

                continue

            sheet = workbook[s_name]

            cfg = HORIZONTAL_RP_GEOMETRY[report_type].get(s_name)
            if not cfg:
                for k, v in HORIZONTAL_RP_GEOMETRY.get(report_type, {}).items():
                    if k.upper() in s_name.upper() or s_name.upper() in k.upper():
                        cfg = v
                        break
            if not cfg:
                if report_type == "CAI":
                    cfg = HORIZONTAL_RP_GEOMETRY["CAI"].get("7.RP-CAI") or HORIZONTAL_RP_GEOMETRY["CAI"].get("0001")
                elif report_type == "FTI":
                    cfg = HORIZONTAL_RP_GEOMETRY["FTI"].get("RP-FTI" if "FTI" in s_name else "0001-1")
                elif report_type in HORIZONTAL_RP_GEOMETRY:
                    cfg = next(iter(HORIZONTAL_RP_GEOMETRY[report_type].values()), None)
            j_row = cfg["joint_row"] if cfg else (10 if "FTI" in s_name else 14)
            data_cols = cfg["data_cols"] if cfg else ([9, 10, 11, 12, 13, 14] if report_type == "CAI" else [10, 11, 12, 13, 14, 15])

            # 1. Update Header Info (Date, Report No, RFI No) & strip external [1]FORM!

            for r in range(1, 8):

                for c in range(1, min(sheet.max_column + 1, 20)):

                    cell = sheet.cell(row=r, column=c)

                    val = cell.value

                    if isinstance(val, str) and "[" in val and "FORM!" in val:
                        if report_type in {"CAI", "SPI"}:
                            if r == 2:
                                cell.value = d_str
                            elif r == 4:
                                cell.value = rep_no
                            elif r in (6, 7):
                                cell.value = rfi_id
                            else:
                                cell.value = None
                        else:
                            cell.value = d_str if c >= 12 else rep_no

            # Explicit standard header locations

            if report_type == "CAI":
                if sheet.cell(row=2, column=11).value is not None or d_str:
                    sheet.cell(row=2, column=11).value = d_str
                if sheet.cell(row=4, column=9).value is not None or rep_no:
                    sheet.cell(row=4, column=9).value = rep_no
                if sheet.cell(row=6, column=10).value is not None:
                    sheet.cell(row=6, column=10).value = rfi_id
            elif report_type == "SPI":
                date_c = None
                for c_cand in range(10, 15):
                    c_obj = sheet.cell(row=2, column=c_cand)
                    if not isinstance(c_obj, MergedCell):
                        date_c = c_cand
                        break
                if date_c and d_str:
                    sheet.cell(row=2, column=date_c).value = d_str
                if rep_no:
                    sheet.cell(row=4, column=9).value = rep_no
                if s_name == "0001-1":
                    if rfi_id:
                        sheet.cell(row=7, column=9).value = rfi_id
                else:
                    if rfi_id:
                        sheet.cell(row=6, column=10).value = rfi_id
            else:
                if sheet.cell(row=2, column=12).value is not None:
                    sheet.cell(row=2, column=12).value = d_str
                if sheet.cell(row=4, column=10).value is not None:
                    sheet.cell(row=4, column=10).value = rep_no
                if sheet.cell(row=6, column=10).value is not None:
                    sheet.cell(row=6, column=10).value = rfi_id

            # 2. Update Station Text and Drawing Numbers

            for r in range(1, min(15, sheet.max_row + 1)):

                for c in range(1, sheet.max_column + 1):

                    cell = sheet.cell(row=r, column=c)

                    if isinstance(cell.value, str):

                        if "TTC15A" in cell.value or "Trạm thi công" in cell.value:

                            cell.value = update_station_text(cell.value, station_code)

                        new_draw, changed = update_drawing_cell_value(cell.value, dynamic_drawing_no, dynamic_weld_map_no)

                        if changed:

                            cell.value = new_draw

            # Apply standard headers for Page 1

            format_standard_rp_header(

                sheet,

                report_type,

                d_str=d_str,

                rep_no=rep_no,

                rfi_id=rfi_id,

                drawing_no=dynamic_drawing_no,

                weld_map_no=dynamic_weld_map_no,

                location_str=location,

                engineer_name=engineer_name_rp,

                page_num=1,

                num_pages=num_pages,

                offset=0,

            )

            # 3. Populate Joints for Page 1

            sheet_capacity = 1 if (report_type == "FTI" and s_name in ["0001-1", "0001"]) else capacity

            for i in range(capacity):

                col = data_cols[i] if i < len(data_cols) else ((9 if report_type == "CAI" else 10) + i)
                cell = sheet.cell(row=j_row, column=col)

                if i < min(total_items, sheet_capacity):

                    cell.value = str(items[i])

                else:

                    cell.value = None

            # 4. Clear Array Formulae so Excel doesn't corrupt the file

            if hasattr(sheet, "array_formulae"):

                sheet.array_formulae.clear()

            # 5. Handle Multi-page Horizontal Pagination

            p_area = sheet.print_area
            if p_area and ":" in p_area:
                m = re.search(r'\$?\w+\$?(\d+)$', p_area)
                s_page_len = int(m.group(1)) if m else (31 if report_type in {"CAI", "FTI", "HTI", "SPI"} else sheet.max_row)
            else:
                s_page_len = 31 if (report_type in {"CAI", "FTI", "HTI", "SPI"} or s_name == f"RP-{report_type}") else sheet.max_row
            if s_name == target_sheet.title:
                target_page_len = s_page_len

            if num_pages > 1 and s_name == target_sheet.title:

                for p in range(2, num_pages + 1):

                    offset = (p - 1) * s_page_len
                    for r in range(1, s_page_len + 1):

                        if r in sheet.row_dimensions:

                            sheet.row_dimensions[r + offset].height = sheet.row_dimensions[r].height

                        for c in range(1, sheet.max_column + 1):

                            src_cell = sheet.cell(row=r, column=c)

                            dst_cell = sheet.cell(row=r + offset, column=c)

                            dst_cell.value = src_cell.value

                            dst_cell.data_type = src_cell.data_type

                            if src_cell.has_style:

                                dst_cell.font = copy.copy(src_cell.font)

                                dst_cell.border = copy.copy(src_cell.border)

                                dst_cell.fill = copy.copy(src_cell.fill)

                                dst_cell.number_format = src_cell.number_format

                                dst_cell.protection = copy.copy(src_cell.protection)

                                dst_cell.alignment = copy.copy(src_cell.alignment)

                            if dst_cell.data_type == "f" and dst_cell.value:

                                val = dst_cell.value.text if hasattr(dst_cell.value, 'text') else str(dst_cell.value)

                                dst_cell.value = Translator(val, origin=src_cell.coordinate).translate_formula(dst_cell.coordinate)

                    # Copy merged ranges

                    for merge in list(sheet.merged_cells.ranges):

                        if merge.min_row >= 1 and merge.max_row <= s_page_len:

                            new_merge = f"{get_column_letter(merge.min_col)}{merge.min_row + offset}:{get_column_letter(merge.max_col)}{merge.max_row + offset}"

                            sheet.merge_cells(new_merge)

                    # Apply standard header in page p

                    format_standard_rp_header(

                        sheet,

                        report_type,

                        d_str=d_str,

                        rep_no=rep_no,

                        rfi_id=rfi_id,

                        drawing_no=dynamic_drawing_no,

                        weld_map_no=dynamic_weld_map_no,

                        location_str=location,

                        engineer_name=engineer_name_rp,

                        page_num=p,

                        num_pages=num_pages,

                        offset=offset,

                    )

                    # Populate joints for page p

                    for i in range(capacity):

                        actual_idx = (p - 1) * capacity + i

                        col = data_cols[i] if i < len(data_cols) else ((9 if report_type == "CAI" else 10) + i)
                        cell = sheet.cell(row=j_row + offset, column=col)

                        if actual_idx < total_items:

                            cell.value = str(items[actual_idx])

                        else:

                            cell.value = None

                    # Page number update: "Trang/Page: {p}"

                    for r in range(1, min(10, s_page_len)):

                        for c in range(1, sheet.max_column + 1):

                            cell = sheet.cell(row=r + offset, column=c)

                            if isinstance(cell.value, str) and ("Trang" in cell.value or "Page" in cell.value):

                                cell.value = re.sub(r"\d+", str(p), cell.value)

                    # Form No: Tờ/sheet 0{p}/0{num_pages}

                    for r in range(s_page_len - 5, s_page_len + 1):

                        for c in range(1, sheet.max_column + 1):

                            cell = sheet.cell(row=r + offset, column=c)

                            if isinstance(cell.value, str) and ("Tờ/sheet" in cell.value or "Tờ" in cell.value):

                                cell.value = re.sub(r"(\d+)/(\d+)", f"{p:02d}/{num_pages:02d}", cell.value)

                # Update Page 1 Tờ/sheet if multi-page

                for r in range(s_page_len - 5, s_page_len + 1):

                    for c in range(1, sheet.max_column + 1):

                        cell = sheet.cell(row=r, column=c)

                        if isinstance(cell.value, str) and ("Tờ/sheet" in cell.value or "Tờ" in cell.value):

                            cell.value = re.sub(r"(\d+)/(\d+)", f"01/{num_pages:02d}", cell.value)

                # Update print area
                if report_type == "CAI":
                    sheet.print_area = f"A1:N{s_page_len * num_pages}"
                    sheet.page_setup.orientation = "portrait"
                    sheet.page_setup.paperSize = 9
                    sheet.page_setup.scale = 73
                    sheet.page_setup.fitToWidth = None
                    sheet.page_setup.fitToHeight = 0
                elif report_type == "FTI" and s_name == target_sheet.title:
                    sheet.print_area = f"A1:O{s_page_len * num_pages}"
                    sheet.page_setup.orientation = "portrait"
                    sheet.page_setup.paperSize = 9
                    sheet.page_setup.scale = 64
                    sheet.page_setup.fitToWidth = 1
                    sheet.page_setup.fitToHeight = 0
                    if not sheet.sheet_properties:
                        from openpyxl.worksheet.properties import WorksheetProperties
                        sheet.sheet_properties = WorksheetProperties()
                    if sheet.sheet_properties.pageSetUpPr is None:
                        from openpyxl.worksheet.properties import PageSetUpPr
                        sheet.sheet_properties.pageSetUpPr = PageSetUpPr()
                    sheet.sheet_properties.pageSetUpPr.fitToPage = True
                    sheet.page_margins.left = 0.7086614173228347
                    sheet.page_margins.right = 0.3149606299212598
                    sheet.page_margins.top = 0.3149606299212598
                    sheet.page_margins.bottom = 0.3149606299212598
                    sheet.page_margins.header = 0.0
                    sheet.page_margins.footer = 0.0
                    sheet.print_options.horizontalCentered = True
                    sheet.print_options.verticalCentered = True
                    if num_pages > 1:
                        from openpyxl.worksheet.pagebreak import Break
                        existing_breaks = {b.id for b in getattr(sheet.row_breaks, "brk", []) if hasattr(b, "id")}
                        for p_brk in range(1, num_pages):
                            brk_row = p_brk * s_page_len
                            if brk_row not in existing_breaks:
                                sheet.row_breaks.append(Break(id=brk_row))
                elif sheet.print_area:
                    parts = sheet.print_area.split(":")
                    if len(parts) == 2:
                        last_part = re.sub(r'\d+$', str(s_page_len * (num_pages if s_name == target_sheet.title else 1)), parts[1])
                        sheet.print_area = f"{parts[0]}:{last_part}"

            # If this is 0001-1 sheet, update its standard header and page number

            if report_type in ["FTI", "SPI"] and s_name == "0001-1":

                start_p = num_pages + 1

                format_standard_rp_header(

                    sheet,

                    report_type,

                    d_str=d_str,

                    rep_no=rep_no,

                    rfi_id=rfi_id,

                    drawing_no=dynamic_drawing_no,

                    weld_map_no=dynamic_weld_map_no,

                    location_str=location,

                    engineer_name=engineer_name_rp,

                    page_num=start_p,

                    num_pages=num_pages,

                    offset=0,

                )

            page_len = target_page_len

# =========================================================================

    # CASE 1.5: MRIR (Material Receiving Inspection Report)

    # =========================================================================

    elif report_type == "MRIR":

        # Giữ nguyên 100% biểu mẫu đầu vào (các hàng vật tư hàn 8..10 và công thức),

        # chỉ thay Report No (T5), Date (V6), Inspection Location (R6), và Họ tên kỹ thuật (D17) nếu có.

        r_data = form_rows.get(rfi_id, {})

        rep_no = str(r_data.get(13) or "").strip()

        if not rep_no:

            rep_no = rfi_id.replace("-RFI-", "-")

        d_val = r_data.get(14) or r_data.get(16) or r_data.get(27)

        if isinstance(d_val, (datetime, date)):

            d_obj = datetime(d_val.year, d_val.month, d_val.day, 0, 0)

        else:

            try:

                d_clean = str(d_val or "").strip()

                d_obj = datetime.strptime(d_clean, "%Y-%m-%d")

            except Exception:

                d_obj = d_val

        red_font_10 = Font(name="Times New Roman", size=10, bold=True, color="FFFF0000")

        red_font_9 = Font(name="Times New Roman", size=9, bold=True, color="FFFF0000")

        # T5: Báo cáo số/Report No.

        target_sheet["T5"].value = rep_no

        target_sheet["T5"].font = red_font_10

        target_sheet["T5"].alignment = Alignment(horizontal="center", vertical="center")

        # V6: Ngày/Date

        target_sheet["V6"].value = d_obj

        target_sheet["V6"].font = red_font_10

        target_sheet["V6"].number_format = '[$-409]dd/mm/yyyy;@'

        target_sheet["V6"].alignment = Alignment(horizontal="left", vertical="center")

        # R6: Vị trí kiểm tra/Inspection Location

        target_sheet["R6"].value = station_code

        target_sheet["R6"].font = red_font_9

        target_sheet["R6"].alignment = Alignment(horizontal="left", vertical="center")

        # D17: Họ và tên Kỹ thuật thi công

        engineer_name_rp = (

            "Nguyễn Đức Sơn"

            if ("TTC15" in location.upper() or "T15" in location.upper())

            else ((site_cfg.get("engineer_name") if site_cfg else None) or default_engineer_for_location(location))

        )

        if engineer_name_rp and str(engineer_name_rp).strip():

            target_sheet["D17"].value = str(engineer_name_rp).strip()

            target_sheet["D17"].font = Font(name="Times New Roman", size=10, bold=False, italic=False)

    # =========================================================================

    # CASE 2A: SPECIALIZED TABULAR HANDLER FOR ICR

    # =========================================================================

    elif report_type == "ICR":

        pipe_items = get_icr_pipe_items(rfi_id, pipe_table)

        capacity = 8

        page_len = 24

        total_items = len(pipe_items)

        num_pages = max(1, (total_items + capacity - 1) // capacity) if total_items > 0 else 1

        d_val = None

        if rfi_id in form_rows:

            d_val = form_rows[rfi_id].get(16) or form_rows[rfi_id].get(14) or form_rows[rfi_id].get(26) or form_rows[rfi_id].get(27) or form_rows[rfi_id].get(3)

        if not d_val:

            d_val = engine.evaluate_cell(target_sheet.title, "L3")

        d_str_clean = format_rp_date_str(d_val)

        rep_val = None

        if rfi_id in form_rows:

            rep_val = form_rows[rfi_id].get(13)

        if not rep_val:

            if target_sheet["K5"].data_type == "f":

                rep_val = engine.evaluate_cell(target_sheet.title, "K5")

        rep_no_clean = str(rep_val or "").strip()

        if not rep_no_clean or ("T15A" in rep_no_clean and "T15A" not in rfi_id):

            rep_no_clean = rfi_id.replace("-RFI-", "-")

        engineer_name_rp = (

            "Nguyễn Đức Sơn"

            if ("TTC15" in location.upper() or "T15" in location.upper())

            else ((site_cfg.get("engineer_name") if site_cfg else None) or default_engineer_for_location(location))

        )

        empty_row_ranges = format_icr_sheet(

            target_sheet,

            d_str=d_str_clean,

            rep_no=rep_no_clean,

            rfi_id=rfi_id,

            drawing_no=dynamic_drawing_no,

            weld_map_no=dynamic_weld_map_no,

            location_str=location,

            engineer_name=engineer_name_rp,

            pipe_items=pipe_items,

            num_pages=num_pages,

            page_len=page_len,

        )

    # =========================================================================

    # CASE 2: VERTICAL TABULAR LAYOUT (FUR, VSR, NDT, etc.)

    # =========================================================================

    elif layout and layout["capacity"] > 0:

        capacity = int(layout["capacity"])

        page_len = 22 if report_type in {"FUR", "MTR"} else int(layout["page_len"])

        t_start = int(layout["table_start"])

        t_end = int(layout["table_end"])

        key_col = layout["key_col"]

        last_tbl_col = layout.get("last_col", target_sheet.max_column)

        key_col_letter = get_column_letter(key_col) if key_col else None

        num_pages = max(1, (total_items + capacity - 1) // capacity) if total_items > 0 else 1

        # Store template formulas from first table row

        row1_formulas: dict[int, str] = {}

        row1_static_values: dict[int, object] = {}

        for c in range(1, target_sheet.max_column + 1):

            t_cell = target_sheet.cell(row=t_start, column=c)

            if t_cell.data_type == "f" and t_cell.value:

                row1_formulas[c] = str(t_cell.value).lstrip("=")

            elif t_cell.value is not None:

                row1_static_values[c] = t_cell.value

        # 1. Evaluate Page 1 cells (rows 1..page_len)

        for r in range(1, page_len + 1):

            for c in range(1, target_sheet.max_column + 1):

                cell = target_sheet.cell(row=r, column=c)

                if isinstance(cell, MergedCell):

                    continue

                coord = cell.coordinate

                if cell.data_type == "f":

                    try:

                        val = engine.evaluate_cell(target_sheet.title, coord)

                        cell.value = val

                        if layout and t_start <= r <= t_end:

                            slash_val = format_rp_multiline_slash(cell.value)

                            if slash_val:

                                cell.value = slash_val

                                curr_a = cell.alignment

                                cell.alignment = Alignment(

                                    horizontal=curr_a.horizontal if curr_a else None,

                                    vertical=curr_a.vertical if curr_a else "center",

                                    wrap_text=True,

                                )

                    except Exception:

                        pass

                elif cell.value is not None:

                    if is_station_cell(cell.value):

                        cell.value = update_station_text(str(cell.value), station_code)

                    new_draw, changed = update_drawing_cell_value(cell.value, dynamic_drawing_no, dynamic_weld_map_no)

                    if changed:

                        cell.value = new_draw

                    if layout and t_start <= r <= t_end:

                        slash_val = format_rp_multiline_slash(cell.value)

                        if slash_val:

                            cell.value = slash_val

                            curr_a = cell.alignment

                            cell.alignment = Alignment(

                                horizontal=curr_a.horizontal if curr_a else None,

                                vertical=curr_a.vertical if curr_a else "center",

                                wrap_text=True,

                            )

        # Fallback row counting if total_items == 0

        if total_items == 0 and layout and layout["capacity"] > 0 and key_col and key_col_letter:

            count = 0

            for r in range(t_start, t_end + 1):

                val = engine.evaluate_cell(target_sheet.title, f"{key_col_letter}{r}")

                if val is not None and str(val).strip() not in {"", "/", "0", "#N/A", "#VALUE!"}:

                    count += 1

            if count > 0:

                total_items = count

                num_pages = max(1, (total_items + capacity - 1) // capacity)

        # Extract Date and Report No for headers

        d_val = None

        date_coords = {"FUR": "N3", "VSR": "N3", "DIC": "R3", "MTR": "L3"}

        if report_type in date_coords:

            try:

                d_val = engine.evaluate_cell(target_sheet.title, date_coords[report_type])

            except Exception:

                d_val = None

        if not d_val and rfi_id in form_rows:

            d_val = form_rows[rfi_id].get(26) or form_rows[rfi_id].get(16) or form_rows[rfi_id].get(14) or form_rows[rfi_id].get(27)

        d_str = format_rp_date_str(d_val)

        rep_val = None

        rep_coords = {"FUR": "M5", "VSR": "M5", "DIC": "O5", "MTR": "K5"}

        if report_type in rep_coords:

            try:

                rep_val = engine.evaluate_cell(target_sheet.title, rep_coords[report_type])

            except Exception:

                rep_val = None

        if not rep_val and rfi_id in form_rows:

            rep_val = form_rows[rfi_id].get(13)

        rep_no = str(rep_val or "").strip()

        if not rep_no:

            rep_no = rfi_id.replace("-RFI-", "-")

        # Format standard header for Page 1

        format_standard_rp_header(

            target_sheet,

            report_type,

            d_str=d_str,

            rep_no=rep_no,

            rfi_id=rfi_id,

            drawing_no=dynamic_drawing_no,

            weld_map_no=dynamic_weld_map_no,

            location_str=location,

            engineer_name=engineer_name_rp,

            page_num=1,

            num_pages=num_pages,

            offset=0,

        )

        # 2. Duplicate pages if num_pages > 1

        if num_pages > 1:

            for p in range(2, num_pages + 1):

                offset = page_len * (p - 1)

                for r in range(1, page_len + 1):

                    if r in target_sheet.row_dimensions:

                        target_sheet.row_dimensions[r + offset].height = target_sheet.row_dimensions[r].height

                    for c in range(1, target_sheet.max_column + 1):

                        src_cell = target_sheet.cell(row=r, column=c)

                        dst_cell = target_sheet.cell(row=r + offset, column=c)

                        dst_cell.value = src_cell.value

                        dst_cell.data_type = src_cell.data_type

                        if src_cell.has_style:

                            dst_cell.font = copy.copy(src_cell.font)

                            dst_cell.border = copy.copy(src_cell.border)

                            dst_cell.fill = copy.copy(src_cell.fill)

                            dst_cell.number_format = src_cell.number_format

                            dst_cell.protection = copy.copy(src_cell.protection)

                            dst_cell.alignment = copy.copy(src_cell.alignment)

                for merge in list(target_sheet.merged_cells.ranges):

                    if merge.min_row >= 1 and merge.max_row <= page_len:

                        new_merge = f"{get_column_letter(merge.min_col)}{merge.min_row + offset}:{get_column_letter(merge.max_col)}{merge.max_row + offset}"

                        target_sheet.merge_cells(new_merge)

                for r in range(1, min(15, page_len)):

                    for c in range(1, 16):

                        cell = target_sheet.cell(row=r + offset, column=c)

                        if isinstance(cell.value, str) and ("Trang" in cell.value or "Page" in cell.value):

                            cell.value = re.sub(r"\d+", str(p), cell.value)

                # Format standard header for page p

                format_standard_rp_header(

                    target_sheet,

                    report_type,

                    d_str=d_str,

                    rep_no=rep_no,

                    rfi_id=rfi_id,

                    drawing_no=dynamic_drawing_no,

                    weld_map_no=dynamic_weld_map_no,

                    location_str=location,

                    engineer_name=engineer_name_rp,

                    page_num=p,

                    num_pages=num_pages,

                    offset=offset,

                )

                # Re-evaluate table rows for page p

                p_t_start = t_start + offset

                p_t_end = t_end + offset

                row_offset = (p - 1) * capacity

                for r in range(p_t_start, p_t_end + 1):

                    for c in range(1, target_sheet.max_column + 1):

                        dst_cell = target_sheet.cell(row=r, column=c)

                        if c in row1_formulas:

                            formula_str = row1_formulas[c]

                            def offset_row_ref(match: re.Match) -> str:

                                col_letters = match.group(1)

                                row_num = int(match.group(2))

                                return f"{col_letters}{row_num + row_offset}"

                            shifted = re.sub(r"([A-Za-z]+)(\d+)", offset_row_ref, formula_str)

                            try:

                                val = engine.evaluate_formula(shifted, target_sheet.title)

                                dst_cell.value = val

                                slash_val = format_rp_multiline_slash(dst_cell.value)

                                if slash_val:

                                    dst_cell.value = slash_val

                                    curr_a = dst_cell.alignment

                                    dst_cell.alignment = Alignment(

                                        horizontal=curr_a.horizontal if curr_a else None,

                                        vertical=curr_a.vertical if curr_a else "center",

                                        wrap_text=True,

                                    )

                            except Exception:

                                pass

                        elif c in row1_static_values:

                            dst_cell.value = row1_static_values[c]

            if target_sheet.print_area:

                parts = target_sheet.print_area.split(":")

                if len(parts) == 2:

                    last_part = re.sub(r'\d+$', str(page_len * num_pages), parts[1])

                    target_sheet.print_area = f"{parts[0]}:{last_part}"

        if report_type == "FUR":

            target_sheet.print_area = f"A1:O{page_len * num_pages}"

            target_sheet.page_setup.orientation = "landscape"

            target_sheet.page_setup.paperSize = 9

            if not target_sheet.sheet_properties:

                from openpyxl.worksheet.properties import WorksheetProperties

                target_sheet.sheet_properties = WorksheetProperties()

            if target_sheet.sheet_properties.pageSetUpPr:

                target_sheet.sheet_properties.pageSetUpPr.fitToPage = True

            target_sheet.page_setup.fitToWidth = 1

            target_sheet.page_setup.fitToHeight = num_pages
        elif report_type == "MTR":
            target_sheet.print_area = f"A1:L{page_len * num_pages}"
            target_sheet.page_setup.orientation = "landscape"
            target_sheet.page_setup.paperSize = 9
            target_sheet.page_setup.scale = 87
            target_sheet.page_setup.fitToWidth = 1
            target_sheet.page_setup.fitToHeight = 0
            if not target_sheet.sheet_properties:
                from openpyxl.worksheet.properties import WorksheetProperties
                target_sheet.sheet_properties = WorksheetProperties()
            if target_sheet.sheet_properties.pageSetUpPr is None:
                from openpyxl.worksheet.properties import PageSetUpPr
                target_sheet.sheet_properties.pageSetUpPr = PageSetUpPr()
            target_sheet.sheet_properties.pageSetUpPr.fitToPage = True
            target_sheet.page_margins.left = 0.25
            target_sheet.page_margins.right = 0.25
            target_sheet.page_margins.top = 0.82
            target_sheet.page_margins.bottom = 0.25
            target_sheet.page_margins.header = 0.25
            target_sheet.page_margins.footer = 0.0
            target_sheet.print_options.horizontalCentered = True
            target_sheet.print_options.verticalCentered = False
            if num_pages > 1:
                from openpyxl.worksheet.pagebreak import Break
                existing_breaks = {b.id for b in getattr(target_sheet.row_breaks, "brk", []) if hasattr(b, "id")}
                for p_brk in range(1, num_pages):
                    brk_row = p_brk * 22
                    if brk_row not in existing_breaks:
                        target_sheet.row_breaks.append(Break(id=brk_row))

        # 3. Clean up empty rows across pages

        key_col = layout.get("key_col")

        if not key_col:

            for cand in [2, 4, 1]:

                if cand <= target_sheet.max_column and target_sheet.cell(row=t_start, column=cand).value is not None:

                    key_col = cand

                    break

        if key_col:

            stt_counter = 0

            for p in range(1, num_pages + 1):

                offset = page_len * (p - 1)

                p_t_start = t_start + offset

                p_t_end = t_end + offset

                empty_start = None

                for r in range(p_t_start, p_t_end + 1):

                    val = target_sheet.cell(row=r, column=key_col).value

                    is_empty = val is None or str(val).strip() in {"", "/", "0", "#N/A", "#VALUE!"}

                    if is_empty:

                        if empty_start is None:

                            empty_start = r

                        for c in range(1, target_sheet.max_column + 1):

                            cell = target_sheet.cell(row=r, column=c)

                            if not isinstance(cell, MergedCell):

                                cell.value = None

                    else:

                        stt_counter += 1

                        if layout.get("is_stt"):

                            target_sheet.cell(row=r, column=1).value = stt_counter

                        empty_start = None

                if empty_start is not None and empty_start <= p_t_end:

                    col_start_0 = 1 if report_type == "FUR" else 0

                    col_end_0 = 13 if report_type == "FUR" else last_tbl_col

                    row_start_0 = empty_start - 1

                    row_end_0 = p_t_end

                    empty_row_ranges.append((col_start_0, row_start_0, col_end_0, row_end_0))

    # =========================================================================

    # CASE 3: FALLBACK EVALUATION

    # =========================================================================

    else:

        for cell in iter_formula_cells(target_sheet):

            try:

                val = engine.evaluate_cell(target_sheet.title, cell.coordinate)

                if isinstance(val, str):

                    val = re.sub(r"TTC15A?", station_code, val)

                cell.value = val

                slash_val = format_rp_multiline_slash(cell.value)

                if slash_val:

                    cell.value = slash_val

                    curr_a = cell.alignment

                    cell.alignment = Alignment(

                        horizontal=curr_a.horizontal if curr_a else None,

                        vertical=curr_a.vertical if curr_a else "center",

                        wrap_text=True,

                    )

            except Exception:

                pass

        if hasattr(target_sheet, "array_formulae"):

            target_sheet.array_formulae.clear()

        for r in range(1, target_sheet.max_row + 1):

            for c in range(1, target_sheet.max_column + 1):

                c_obj = target_sheet.cell(row=r, column=c)

                if c_obj.data_type == "s" and isinstance(c_obj.value, str):

                    if "TTC15" in c_obj.value:

                        c_obj.value = re.sub(r"TTC15A?", station_code, c_obj.value)

    # Evaluate any formula cells in other sheets to avoid broken external links

    for sheet in workbook.worksheets:

        if sheet.title == target_sheet.title or sheet.title.casefold().startswith("controlog") or sheet.title == "MRIR":

            continue

        if report_type in HORIZONTAL_RP_GEOMETRY and sheet.title in ["0001-1", "0001"]:

            continue  # Already updated above

        for cell in iter_formula_cells(sheet):

            try:

                val = engine.evaluate_cell(sheet.title, cell.coordinate)

                if isinstance(val, str):

                    val = re.sub(r"TTC15A?", station_code, val)

                cell.value = val

            except Exception:

                pass

        if hasattr(sheet, "array_formulae"):

            sheet.array_formulae.clear()

        for r in range(1, sheet.max_row + 1):

            for c in range(1, sheet.max_column + 1):

                c_obj = sheet.cell(row=r, column=c)

                if c_obj.data_type == "s" and isinstance(c_obj.value, str):

                    if "TTC15" in c_obj.value:

                        c_obj.value = re.sub(r"TTC15A?", station_code, c_obj.value)

    # Replace Area with Location

    for sheet in workbook.worksheets:

        if sheet.title == "MRIR" or (report_type == "ICR" and sheet.title == target_sheet.title):

            continue

        replace_area_with_location(sheet)

    # Normalize fonts and signatures

    engineer_name_rp = (

        "Nguyễn Đức Sơn"

        if ("TTC15" in location.upper() or "T15" in location.upper())

        else ((site_cfg.get("engineer_name") if site_cfg else None) or default_engineer_for_location(location))

    )

    for sheet in workbook.worksheets:

        if report_type in {"ICR", "FUR"} and sheet.title == target_sheet.title:

            continue

        max_row_limit = (page_len * num_pages) if (sheet.title == target_sheet.title and num_pages > 1) else None

        normalize_sheet_fonts_and_signature(sheet, engineer_name=engineer_name_rp, max_row_limit=max_row_limit)

    if report_type == "SPI":

        for sheet in workbook.worksheets:

            if not sheet.title.casefold().startswith("controlog"):

                thin_s = Side(style="thin", color="000000")

                for r in range(1, sheet.max_row + 1):

                    c1_val = str(sheet.cell(row=r, column=1).value or "")

                    # Ensure vertical border between Item (col 7 / G) and RFI (col 8 / H) on row 6

                    if "địa điểm" in c1_val.lower() or "location" in c1_val.lower():

                        cg = sheet.cell(row=r, column=7)

                        ch = sheet.cell(row=r, column=8)

                        cg.border = Border(

                            left=cg.border.left if cg.border else None,

                            right=thin_s,

                            top=cg.border.top if cg.border else thin_s,

                            bottom=cg.border.bottom if cg.border else thin_s,

                        )

                        ch.border = Border(

                            left=thin_s,

                            right=ch.border.right if ch.border else None,

                            top=ch.border.top if ch.border else thin_s,

                            bottom=ch.border.bottom if ch.border else thin_s,

                        )

                    # Convert all °C in Relative Humidity (Độ ẩm tương đối) row to %

                    if "độ ẩm" in c1_val.lower() or "relative humidity" in c1_val.lower():

                        for c in range(1, sheet.max_column + 1):

                            cell = sheet.cell(row=r, column=c)

                            nf = cell.number_format or ""

                            if "°c" in nf.lower() or "oc" in nf.lower():

                                cell.number_format = '0.0"%"'

                            if isinstance(cell.value, str) and ("°c" in cell.value.lower() or "oc" in cell.value.lower()):

                                cell.value = cell.value.replace("°C", "%").replace("oC", "%").replace("°c", "%")

                        for c in range(8, 14):

                            cell = sheet.cell(row=r, column=c)

                            if cell.number_format in [None, 'General', '', '0.0"°C"', '0.0" °C"']:

                                cell.number_format = '0.0"%"'

    # Enforce standard page margins for all report sheets (L=18mm, R=8mm, T=8mm, B=8mm, H=0, F=0)
    if report_type not in {"SPI", "ICR", "FUR", "VSR", "DIC", "CAI", "MTR", "PPP"}:
        for sheet in workbook.worksheets:

            if sheet.title.casefold().startswith("controlog"):

                continue

            sheet.page_margins.left = 18.0 / 25.4

            sheet.page_margins.right = 8.0 / 25.4

            sheet.page_margins.top = 8.0 / 25.4

            sheet.page_margins.bottom = 8.0 / 25.4

            sheet.page_margins.header = 0.0

            sheet.page_margins.footer = 0.0
    if report_type == "CAI":
        for sheet in workbook.worksheets:
            if sheet.title.casefold().startswith("controlog"):
                continue
            sheet.page_margins.left = 15.0 / 25.4
            sheet.page_margins.right = 5.0 / 25.4
            sheet.page_margins.top = 8.0 / 25.4
            sheet.page_margins.bottom = 8.0 / 25.4
            sheet.page_margins.header = 0.0
            sheet.page_margins.footer = 0.0
            sheet.print_area = f"A1:N{31 * num_pages}"
            sheet.page_setup.orientation = "portrait"
            sheet.page_setup.paperSize = 9
            sheet.page_setup.scale = 73
            sheet.page_setup.fitToWidth = None
            sheet.page_setup.fitToHeight = 0

    report_no = str(form_rows.get(rfi_id, {}).get(13) or "").strip()

    safe_report = re.sub(r"[^A-Za-z0-9._-]+", "_", report_no) if report_no else ""

    safe_id = re.sub(r"[^A-Za-z0-9._-]+", "_", rfi_id)

    file_base = safe_report if safe_report else f"{safe_id}_RP-{report_type}"

    output_path = site_output_dir(location) / f"{file_base}{template.suffix.lower()}"

    temporary_path = output_path.with_name(f"{output_path.name}.{os.getpid()}_{threading.get_ident()}.tmp")

    sheet_joint_counts: dict[str, int] = {}
    for sheet in workbook.worksheets:
        if report_type == "FTI" and sheet.title in ["0001-1", "0001"]:
            sheet_joint_counts[sheet.title] = 1 if total_items > 0 else 0
        elif sheet.title == "MRIR":
            sheet_joint_counts[sheet.title] = 0
        elif sheet.title == target_sheet.title:
            sheet_joint_counts[sheet.title] = total_items
        else:
            cfg_sheet = HORIZONTAL_RP_GEOMETRY.get(report_type, {}).get(sheet.title)
            if not cfg_sheet and report_type in HORIZONTAL_RP_GEOMETRY:
                for k, v in HORIZONTAL_RP_GEOMETRY[report_type].items():
                    if k.upper() in sheet.title.upper() or sheet.title.upper() in k.upper():
                        cfg_sheet = v
                        break
                if not cfg_sheet:
                    cfg_sheet = next(iter(HORIZONTAL_RP_GEOMETRY[report_type].values()), None)
            if cfg_sheet and "joint_row" in cfg_sheet and "data_cols" in cfg_sheet:
                j_r = cfg_sheet["joint_row"]
                act_count = 0
                for c in cfg_sheet["data_cols"]:
                    val = sheet.cell(row=j_r, column=c).value
                    if val is not None and str(val).strip() and not str(val).strip().startswith("="):
                        act_count += 1
                sheet_joint_counts[sheet.title] = act_count if act_count > 0 else total_items
            else:
                sheet_joint_counts[sheet.title] = total_items

    # Dedicated formatting for PPP

    if report_type == "PPP":

        d_val = (

            form_rows.get(rfi_id, {}).get(16)

            or form_rows.get(rfi_id, {}).get(14)

            or form_rows.get(rfi_id, {}).get(27)

            or target_sheet["J3"].value

        )

        if isinstance(d_val, (datetime, date)):

            d_str_clean = d_val.strftime("%d/%m/%Y")

        elif d_val:

            d_clean = str(d_val).split()[0].replace("-", "/")

            parts = d_clean.split("/")

            if len(parts) == 3 and len(parts[0]) == 4:

                d_str_clean = f"{parts[2]}/{parts[1]}/{parts[0]}"

            else:

                d_str_clean = d_clean

        else:

            d_str_clean = ""

        rep_no_clean = str(form_rows.get(rfi_id, {}).get(13) or "").strip()

        if not rep_no_clean:

            rep_no_clean = str(target_sheet["J4"].value or "").strip()

        if not rep_no_clean:

            rep_no_clean = rfi_id.replace("-RFI-", "-")

        format_ppp_sheet(

            target_sheet,

            d_str=d_str_clean,

            rep_no=rep_no_clean,

            rfi_id=rfi_id,

            drawing_no=dynamic_drawing_no,

            weld_map_no=dynamic_weld_map_no,

            location_str=location,

            engineer_name=engineer_name_rp,

            num_pages=num_pages,

            page_len=page_len,

        )

    # Format 23°C, Rich text for signatures, and tabSelected

    bold_font = InlineFont(b=True)

    italic_font = InlineFont(i=True)

    sig_kws = ['TỔNG THẦU', 'GIÁM SÁT', 'CHỦ ĐẦU TƯ', 'ĐĂNG KIỂM', 'KỸ THUẬT', 'KỸ THUẬT', 'CHỨC DANH', 'KIỂM TRA', 'REPRESENTATIVE', 'POSITION', 'INSPECTION', 'CHỮ KÝ', 'HỌ VÀ TÊN', 'ĐẠI DIỆN', 'SIGNATURE', 'FULL NAME']

    for sheet in workbook.worksheets:

        if sheet.title == "MRIR" or (report_type in {"PPP", "ICR", "FUR"} and sheet.title == target_sheet.title):

            if sheet.views.sheetView:

                sheet.views.sheetView[0].tabSelected = (sheet.title == target_sheet.title)

            continue

        if sheet.views.sheetView:

            if sheet.title in [target_sheet.title, "0001-1", "0001"]:

                sheet.views.sheetView[0].tabSelected = True

            else:

                sheet.views.sheetView[0].tabSelected = False

        for r in range(1, sheet.max_row + 1):

            for c in range(1, sheet.max_column + 1):

                cell = sheet.cell(row=r, column=c)

                val = cell.value

                if isinstance(val, str):

                    if "23oC" in val:

                        val = val.replace("23oC", "23°C")

                        cell.value = val

                    if ("/" in val or "\n" in val) and any(kw in val.upper() for kw in sig_kws):

                        if "ACC" in val.upper() or "REJ" in val.upper() or "N/A" in val.upper() or "SƠN" in val.upper() or "TRƯỜNG" in val.upper():

                            continue

                        if "\n" in val:

                            parts = val.split("\n", 1)

                            parts[0] += "\n"

                        else:

                            parts = val.split("/", 1)

                            parts[0] += "/"

                        try:

                            cell.value = CellRichText(

                                TextBlock(bold_font, parts[0]),

                                TextBlock(italic_font, parts[1])

                            )

                        except Exception:

                            pass

    # Keep fonts / font sizes and print layout exactly as defined in the input RP template

    target_names = [target_sheet.title]

    if "0001-1" in workbook.sheetnames and "0001-1" not in target_names:

        target_names.append("0001-1")

    restore_template_formatting(

        workbook,

        template,

        target_names,

        page_len=page_len,

        num_pages=num_pages,

        paged_sheet=target_sheet.title,
        report_type=report_type,

    )

    # Copy title / header / footer rich text runs 1:1 from the input template

    for s_name in target_names:

        if s_name in workbook.sheetnames:

            apply_template_rich_text(

                workbook[s_name],

                template,

                page_len=page_len,

                num_pages=num_pages if s_name == target_sheet.title else 1,

                station_code=station_code,

                drawing_no=dynamic_drawing_no,

                weld_map_no=dynamic_weld_map_no,

            )

    standardize_all_rp_titles_before_save(workbook, report_type, page_len=page_len, num_pages=num_pages, paged_sheet_title=target_sheet.title)
    try:
        workbook.save(temporary_path)

    finally:

        workbook.close()

    # Read pristine drawings, rels, and media directly from original template

    with zipfile.ZipFile(template, "r") as tmpl_archive:

        drawings = sheet_drawing_paths(tmpl_archive)

        all_tmpl_media = {

            f: tmpl_archive.read(f)

            for f in tmpl_archive.namelist()

            if f.startswith("xl/media/")

        }

        all_tmpl_rels = {

            f: tmpl_archive.read(f)

            for f in tmpl_archive.namelist()

            if f.startswith("xl/drawings/_rels/")

        }

        tmpl_drawings = {

            f: tmpl_archive.read(f)

            for f in tmpl_archive.namelist()

            if f.startswith("xl/drawings/") and not f.startswith("xl/drawings/_rels/")

        }

    target_drawing = drawings.get(target_sheet.title)

    drawing_to_sheet = {d.lstrip("/"): s for s, d in drawings.items()}

    # SPI sheets preserve original template drawings, rels, and media verbatim
    if report_type == "SPI":
        pass

    # Clean XML and duplicate drawing anchors for multi-page

    cleaned_path = output_path.with_name(output_path.name + ".cleaned")

    try:

        with zipfile.ZipFile(temporary_path, "r") as original:

            with zipfile.ZipFile(

                cleaned_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6

            ) as output:

                written_files: set[str] = set()

                for item in original.infolist():

                    if (

                        item.filename.startswith("xl/externalLinks/")

                        or item.filename == "xl/calcChain.xml"

                        or item.filename.startswith("xl/media/")

                        or item.filename.startswith("xl/drawings/_rels/")

                    ):

                        continue

                    data = original.read(item.filename)

                    if item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                        xml_text = data.decode("utf-8")
                        def _ensure_xml_space(match):
                            full_tag = match.group(0)
                            if 'xml:space="preserve"' in full_tag:
                                return full_tag
                            text_content = match.group(1)
                            if "\n" in text_content or "\r" in text_content or text_content.endswith(" ") or text_content.startswith(" "):
                                return f'<t xml:space="preserve">{text_content}</t>'
                            return full_tag
                        xml_text = re.sub(r'<t(?:\s+[^>]*)?>(.*?)</t>', _ensure_xml_space, xml_text, flags=re.DOTALL)
                        if report_type not in {"SPI", "ICR", "FUR", "VSR", "DIC", "CAI", "MTR", "PPP"}:
                            xml_text = re.sub(
                                r'<pageMargins\b[^>]*/>',
                                '<pageMargins left="0.7086614173228347" right="0.3149606299212598" top="0.3149606299212598" bottom="0.3149606299212598" header="0" footer="0"/>',
                                xml_text,
                            )
                        elif report_type == "CAI":
                            xml_text = re.sub(
                                r'<pageMargins\b[^>]*/>',
                                '<pageMargins left="0.5905511811023623" right="0.1968503937007874" top="0.31496062992125984" bottom="0.31496062992125984" header="0" footer="0"/>',
                                xml_text,
                            )
                        elif report_type == "PPP":
                            xml_text = re.sub(
                                r'<pageMargins\b[^>]*/>',
                                '<pageMargins left="0.27" right="0.16" top="0.62" bottom="0.2" header="0" footer="0"/>',
                                xml_text,
                            )
                        data = xml_text.encode("utf-8")
                    elif item.filename == "xl/workbook.xml":

                        data = clean_workbook_xml(data)

                    elif item.filename == "xl/styles.xml":

                        data = clean_styles_xml(data, preserve_red=(report_type == "MRIR"))

                    elif item.filename == "xl/_rels/workbook.xml.rels":

                        data = re.sub(

                            rb'<Relationship\b[^>]*Type="[^"]*(?:externalLink|calcChain)[^"]*"[^>]*/>',

                            b"",

                            data,

                        )

                    elif item.filename == "[Content_Types].xml":

                        data = re.sub(

                            rb'<Override\b[^>]*PartName="[^"]*(?:externalLink|calcChain)[^"]*"[^>]*/>',

                            b"",

                            data,

                        )

                        for ext in [b"png", b"jpeg", b"jpg", b"emf"]:

                            if b'Extension="' + ext + b'"' not in data:

                                ct_type = b"image/" + (b"x-emf" if ext == b"emf" else ext)

                                data = data.replace(

                                    b"</Types>",

                                    rb'<Default Extension="' + ext + rb'" ContentType="' + ct_type + rb'"/></Types>',

                                )

                    elif item.filename.lstrip("/") in drawing_to_sheet or item.filename in tmpl_drawings:

                        sheet_name_for_drawing = drawing_to_sheet.get(item.filename.lstrip("/"), "")

                        base_drawing_xml = tmpl_drawings.get(

                            item.filename.lstrip("/"), tmpl_drawings.get(item.filename, data)

                        )

                        if report_type in HORIZONTAL_RP_GEOMETRY:

                            j_count = sheet_joint_counts.get(sheet_name_for_drawing, total_items)

                            p_pages = num_pages if sheet_name_for_drawing == target_sheet.title else 1

                            data = apply_horizontal_diagonal_line(

                                base_drawing_xml,

                                report_type,

                                sheet_name_for_drawing,

                                j_count,

                                page_len=page_len,

                                num_pages=p_pages,

                            )

                        else:

                            data = adjust_drawing_lines(base_drawing_xml, sheet_name_for_drawing, None, {})

                            if target_drawing and item.filename.lstrip("/") == target_drawing.lstrip("/"):

                                data = update_drawing_anchors_and_locks(

                                    data, num_pages, page_len, empty_row_ranges

                                )

                    output.writestr(item, data)

                    written_files.add(item.filename)

                # Write any drawings from template not present in original

                for d_name, d_data in tmpl_drawings.items():

                    if d_name not in written_files:

                        output.writestr(d_name, d_data)

                # Write all media files and drawing rels from the pristine template

                for m_name, m_data in all_tmpl_media.items():

                    output.writestr(m_name, m_data)

                for r_name, r_data in all_tmpl_rels.items():

                    output.writestr(r_name, r_data)

        cleaned_path.replace(output_path)

    finally:

        if temporary_path.exists():

            temporary_path.unlink()

        if cleaned_path.exists():

            cleaned_path.unlink()

    return output_path, report_type

def generate_workbook(

    rfi_id: str,

    form_rows: dict[str, dict[int, object]],

    location: str = DEFAULT_SITE,

    external_tables: dict[str, list[tuple[int, tuple[object, ...]]]] | None = None,

) -> tuple[Path, str]:

    location = validate_site_code(location, require_existing=True)

    report_type = get_rfi_type(rfi_id)

    template = find_template(report_type or "", location)

    if template is None:

        raise UserError(f"Chưa có biểu mẫu RFI cho loại {report_type or 'không xác định'}.")

    source_row = form_rows.get(rfi_id)

    if source_row is None:

        raise UserError(f"Không tìm thấy số RFI {rfi_id} trong dữ liệu hiện tại.")

    try:

        workbook = load_workbook(

            template,

            read_only=False,

            data_only=False,

            keep_vba=template.suffix.lower() == ".xlsm",

            keep_links=True,

        )

    except Exception as exc:

        raise UserError(f"Không đọc được biểu mẫu {template.name}: {exc}") from exc

    formula_values: dict[tuple[str, str], object] = {}

    selector_sheet: str | None = None

    station_cell_updates: dict[tuple[str, str], str] = {}

    drawing_cell_updates: dict[tuple[str, str], str] = {}

    station_code = station_code_for_location(location)

    site_cfg = get_site_config(location)

    if external_tables is None:

        try:

            external_tables = get_cached_rp_source_tables(location)

        except Exception:

            external_tables = {}

    dynamic_drawing_no, dynamic_weld_map_no = resolve_drawings_for_rfi(

        rfi_id,

        report_type or "",

        external_tables,

        location,

        fallback_drawing=site_cfg["drawing_no"],

        fallback_weld_map=site_cfg["weld_map_no"],

    )

    if report_type in {"ICR", "DIC"}:

        dynamic_drawing_no = re.sub(r"\s+(?:SHT|Sheet).*$", "", dynamic_drawing_no, flags=re.IGNORECASE).strip()

        dynamic_weld_map_no = re.sub(r"\s+(?:SHT|Sheet).*$", "", dynamic_weld_map_no, flags=re.IGNORECASE).strip()

    engineer_name_rfi = site_cfg.get("engineer_name", default_engineer_for_location(location))

    engineer_cell_updates: dict[tuple[str, str], str] = {}

    for sheet in workbook.worksheets:

        selector = sheet["A6"]

        if selector.value is not None and not (selector.data_type == "f"):

            if "RFI-" in str(selector.value).upper():

                selector_sheet = sheet.title

        for row in sheet.iter_rows():

            for cell in row:

                if cell.data_type == "f":

                    formula = formula_text(cell.value)

                    if "[" not in formula:

                        raise UserError(

                            f"Biểu mẫu {template.name} có công thức chưa được hỗ trợ "

                            f"tại {sheet.title}!{cell.coordinate}."

                        )

                    formula_values[(sheet.title, cell.coordinate)] = evaluate_form_formula(

                        formula, rfi_id, form_rows

                    )

                elif is_station_cell(cell.value):

                    station_cell_updates[(sheet.title, cell.coordinate)] = update_station_text(

                        str(cell.value), station_code

                    )

                else:

                    curr_val = cell.value

                    if isinstance(curr_val, str) and ("Nguyễn Đức Sơn" in curr_val or "Lê Đình Trường" in curr_val):

                        if engineer_name_rfi:

                            engineer_cell_updates[(sheet.title, cell.coordinate)] = engineer_name_rfi

                    if isinstance(curr_val, str) and ("khu vực" in curr_val.lower() or "area" in curr_val.lower()):

                        new_loc = re.sub(

                            r"Khu\s*v[ựu]c\s*/\s*Area\s*:\s*",

                            "Địa điểm/ Location: ",

                            curr_val,

                            flags=re.IGNORECASE,

                        )

                        if new_loc == curr_val:

                            new_loc = re.sub(

                                r"Khu\s*v[ựu]c\s*/\s*Area\b",

                                "Địa điểm/ Location",

                                curr_val,

                                flags=re.IGNORECASE,

                            )

                        if new_loc != curr_val:

                            curr_val = new_loc

                            drawing_cell_updates[(sheet.title, cell.coordinate)] = curr_val

                    new_drawing_val, drawing_changed = update_drawing_cell_value(

                        curr_val, dynamic_drawing_no, dynamic_weld_map_no

                    )

                    if drawing_changed:

                        drawing_cell_updates[(sheet.title, cell.coordinate)] = new_drawing_val

    if not formula_values:

        raise UserError(f"Biểu mẫu {template.name} không có công thức liên kết cần cập nhật.")

    if selector_sheet is None:

        selector_sheet = workbook.worksheets[0].title

    with zipfile.ZipFile(template, "r") as template_archive:

        paths = worksheet_paths(template_archive)

    workbook.close()

    safe_id = re.sub(r"[^A-Za-z0-9._-]+", "_", rfi_id)

    output_name = f"{safe_id}_{report_type}{template.suffix.lower()}"

    output_path = site_output_dir(location) / output_name

    temporary_path = output_path.with_name(f"{output_path.name}.{os.getpid()}_{threading.get_ident()}.tmp")

    try:

        with zipfile.ZipFile(template, "r") as original:

            sheet_updates: dict[str, bytes] = {}

            for (sheet_name, coordinate), value in formula_values.items():

                part = paths.get(sheet_name)

                if part is None:

                    raise UserError(f"Không tìm thấy sheet XML của {sheet_name}.")

                if part not in sheet_updates:

                    sheet_updates[part] = original.read(part)

                sheet_updates[part] = replace_cell_content(sheet_updates[part], coordinate, value)

            selector_part = paths[selector_sheet]

            if selector_part not in sheet_updates:

                sheet_updates[selector_part] = original.read(selector_part)

            sheet_updates[selector_part] = replace_cell_content(

                sheet_updates[selector_part], "A6", rfi_id

            )

            for (sheet_name, coordinate), text_value in station_cell_updates.items():

                part = paths.get(sheet_name)

                if part is None:

                    continue

                if part not in sheet_updates:

                    sheet_updates[part] = original.read(part)

                sheet_updates[part] = replace_cell_content(

                    sheet_updates[part], coordinate, text_value

                )

            for (sheet_name, coordinate), text_value in drawing_cell_updates.items():

                part = paths.get(sheet_name)

                if part is None:

                    continue

                if part not in sheet_updates:

                    sheet_updates[part] = original.read(part)

                sheet_updates[part] = replace_cell_content(

                    sheet_updates[part], coordinate, text_value

                )

            for (sheet_name, coordinate), text_value in engineer_cell_updates.items():

                part = paths.get(sheet_name)

                if part is None:

                    continue

                if part not in sheet_updates:

                    sheet_updates[part] = original.read(part)

                sheet_updates[part] = replace_cell_content(

                    sheet_updates[part], coordinate, text_value

                )

            with zipfile.ZipFile(

                temporary_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=1

            ) as output:

                for item in original.infolist():

                    if item.filename == "xl/calcChain.xml":

                        continue

                    data = sheet_updates.get(item.filename)

                    if data is None:

                        data = original.read(item.filename)

                    if item.filename == "xl/workbook.xml":

                        data = clean_workbook_xml(data)

                    elif item.filename == "xl/styles.xml":

                        data = clean_styles_xml(data)

                    elif item.filename == "xl/_rels/workbook.xml.rels":

                        data = re.sub(

                            rb'<Relationship\b[^>]*Type="[^"]*calcChain[^"]*"[^>]*/>',

                            b"",

                            data,

                        )

                    elif item.filename == "[Content_Types].xml":

                        data = re.sub(

                            rb'<Override\b[^>]*PartName="[^"]*calcChain[^"]*"[^>]*/>',

                            b"",

                            data,

                        )

                    output.writestr(item, data)

        temporary_path.replace(output_path)

    finally:

        if temporary_path.exists():

            temporary_path.unlink()

    return output_path, report_type or ""

def get_site_drawing_path(location: str) -> Path | None:
    location = validate_site_code(location, require_existing=False)
    data_dir = site_data_dir(location)
    cand = data_dir / f"Ban ve {location}.pdf"
    if cand.is_file():
        return cand
    if data_dir.is_dir():
        for item in data_dir.iterdir():
            if item.is_file() and item.suffix.lower() == ".pdf" and not item.name.startswith("~"):
                return item
    fb = Path(f"C:/Users/LE TRUONG/Desktop/IN RFI-RP-DW/INPUT/Ban ve {location}.pdf")
    if fb.is_file():
        return fb
    return None


def get_site_drawing_info(location: str) -> dict[str, object]:
    location = validate_site_code(location, require_existing=False)
    p = get_site_drawing_path(location)
    if not p or not p.is_file():
        return {
            "has_drawing": False,
            "filename": None,
            "message": f"Chưa có file bản vẽ CAD cho trạm {location}"
        }
    try:
        import pymupdf
        doc = pymupdf.open(str(p))
        pages = len(doc)
        doc.close()
    except Exception:
        pages = 1
    st = p.stat()
    sz_kb = round(st.st_size / 1024, 1)
    sz_str = f"{sz_kb} KB" if sz_kb < 1024 else f"{round(sz_kb / 1024, 2)} MB"
    return {
        "has_drawing": True,
        "filename": p.name,
        "pages": pages,
        "size": sz_str,
        "updated": datetime.fromtimestamp(st.st_mtime).strftime("%d/%m/%Y %H:%M"),
        "path": str(p)
    }


def save_uploaded_drawing(location: str, filename: str, content: bytes) -> dict[str, object]:
    location = validate_site_code(location, require_existing=True)
    if not filename.lower().endswith(".pdf"):
        raise UserError("File bản vẽ phải ở định dạng PDF (.pdf).")
    try:
        import pymupdf
        doc = pymupdf.open(stream=content, filetype="pdf")
        page_count = len(doc)
        doc.close()
    except Exception as exc:
        raise UserError(f"File PDF bản vẽ không hợp lệ hoặc bị lỗi: {exc}")
    data_dir = site_data_dir(location)
    data_dir.mkdir(parents=True, exist_ok=True)
    clean_name = f"Ban ve {location}.pdf"
    target_path = data_dir / clean_name
    if target_path.is_file():
        bak_dir = data_dir / "Backup"
        bak_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            shutil.copy2(target_path, bak_dir / f"{ts}_{clean_name}")
        except Exception:
            pass
    target_path.write_bytes(content)
    try:
        cfg = get_site_config(location)
        cfg["drawing_pdf_file"] = clean_name
        save_site_config(location, cfg)
    except Exception:
        pass
    return get_site_drawing_info(location)


def create_combined_pdf_download(
    unique_ids: list[str],
    location: str,
    form_rows: dict[str, dict[int, object]],
    external_tables: dict[str, list[tuple[int, tuple[object, ...]]]],
    site_cfg: dict[str, str] | None = None
) -> tuple[Path, str]:
    import core_processor
    dw_path = get_site_drawing_path(location)
    out_dir = site_output_dir(location)
    out_dir.mkdir(parents=True, exist_ok=True)
    
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp_in:
        for r_id in unique_ids:
            try:
                rfi_p, _ = generate_workbook(r_id, form_rows, location, external_tables)
                shutil.copy(rfi_p, tmp_in)
            except Exception as e:
                LOGGER.warning("Could not generate RFI %s: %s", r_id, e)
            try:
                rp_p, _ = generate_rp_workbook(r_id, form_rows, external_tables, location, site_cfg=site_cfg)
                shutil.copy(rp_p, tmp_in)
            except Exception as e:
                LOGGER.warning("Could not generate RP for %s: %s", r_id, e)
                
        dw_str = str(dw_path.resolve()) if dw_path and dw_path.is_file() else None
        res = core_processor.batch_process(
            input_dir=tmp_in,
            output_dir=str(out_dir.resolve()),
            drawing_pdf_path=dw_str
        )
        if not res.get("success"):
            msg = res.get("message", "Lỗi khi xử lý gộp PDF.")
            if "Không tìm thấy file bản vẽ" in msg:
                msg = f"Chưa có file Bản vẽ CAD (Weld Map PDF) cho trạm {location}. Vui lòng nạp file Bản vẽ tại Tab 'Nhập biểu mẫu' trước khi xuất PDF."
            raise UserError(msg)
            
        comb_pdf = Path(res["combined_pdf"])
        if not comb_pdf.is_file():
            raise UserError("Không tìm thấy file PDF tổng được tạo.")
        return comb_pdf, comb_pdf.name

def create_download(

    selected_ids: list[str], mode: str = "rfi", location: str = DEFAULT_SITE

) -> tuple[Path, str]:

    location = validate_site_code(location, require_existing=True)

    records, form_rows, _dup_ids = load_records(location)

    if mode not in {"rfi", "rp", "pdf"}:

        raise UserError("Không nhận diện được loại biểu mẫu cần xuất.")

    support_key = "supported" if mode in ("rfi", "pdf") else "rp_supported"

    known_ids = {record["id"] for record in records if record[support_key]}

    report_to_id = {record["report"]: record["id"] for record in records if record.get("report")}

    mapped_ids = []

    for item in selected_ids:

        if item in known_ids:

            mapped_ids.append(item)

        elif item in report_to_id and report_to_id[item] in known_ids:

            mapped_ids.append(report_to_id[item])

        else:

            mapped_ids.append(item)

    unique_ids = list(dict.fromkeys(mapped_ids))

    if not unique_ids:

        raise UserError("Hãy chọn ít nhất một biểu mẫu cần xuất.")

    unknown = [rfi_id for rfi_id in unique_ids if rfi_id not in known_ids]

    if unknown:

        raise UserError(

            "Một số mục đã chọn không có biểu mẫu tương ứng trong chế độ đã chọn "

            "hoặc không tồn tại trong dữ liệu."

        )

    external_tables = get_cached_rp_source_tables(location)

    site_cfg = get_site_config(location)

    if mode == "pdf":
        return create_combined_pdf_download(unique_ids, location, form_rows, external_tables, site_cfg)

    worker_count = min(4, max(1, os.cpu_count() or 2))

    if mode == "rfi":

        if len(unique_ids) > 1 and worker_count > 1:

            def _gen_rfi(r_id: str) -> Path:

                return generate_workbook(r_id, form_rows, location, external_tables)[0]

            with ThreadPoolExecutor(max_workers=worker_count) as pool:

                generated = list(pool.map(_gen_rfi, unique_ids))

        else:

            generated = [

                generate_workbook(rfi_id, form_rows, location, external_tables)[0] for rfi_id in unique_ids

            ]

    else:

        if len(unique_ids) > 1 and worker_count > 1:

            def _gen_rp(r_id: str) -> Path:

                return generate_rp_workbook(r_id, form_rows, external_tables, location, site_cfg=site_cfg)[0]

            with ThreadPoolExecutor(max_workers=worker_count) as pool:

                generated = list(pool.map(_gen_rp, unique_ids))

        else:

            generated = [

                generate_rp_workbook(rfi_id, form_rows, external_tables, location, site_cfg=site_cfg)[0]

                for rfi_id in unique_ids

            ]

    if len(generated) == 1:

        return generated[0], generated[0].name

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    zip_path = site_output_dir(location) / f"FOME_{mode.upper()}_{timestamp}.zip"

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=1) as archive:

        for workbook_path in generated:

            archive.write(workbook_path, arcname=workbook_path.name)

    return zip_path, zip_path.name

def backup_current_source(key: str, location: str = DEFAULT_SITE) -> None:

    current = source_path(key, location)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

    backup = site_data_dir(location) / "Backup" / f"{timestamp}_{current.name}"

    backup.parent.mkdir(parents=True, exist_ok=True)

    shutil.copy2(current, backup)

def validate_uploaded_source(

    key: str, content: bytes, location: str = DEFAULT_SITE

) -> None:

    if len(content) > MAX_UPLOAD_BYTES:

        raise UserError("File tải lên vượt giới hạn 100 MB.")

    try:

        workbook = load_workbook(

            __import__("io").BytesIO(content), read_only=True, data_only=True

        )

    except Exception as exc:

        raise UserError(f"File Excel không hợp lệ hoặc không đọc được: {exc}") from exc

    if key == "controlog":

        normalized_names = [normalized_sheet_name(name) for name in workbook.sheetnames]

        has_ctrl = any(

            re.sub(r"[^a-z0-9]", "", name).startswith(("controllog", "controlog"))

            for name in normalized_names

        )

        if not has_ctrl:

            workbook.close()

            raise UserError("File Controlog phải có sheet 'Control Log' (hoặc 'Controlog').")

        has_pipe = any(

            re.sub(r"[^a-z0-9]", "", name).startswith("pipe")

            for name in normalized_names

        )

        if not has_pipe:

            workbook.close()

            raise UserError("File Controlog phải có sheet 'Pipe'.")

    if key == "rfi_data":

        sheet_name = form_sheet_name(workbook)

        if sheet_name is None:

            workbook.close()

            raise UserError("Không tìm thấy sheet FORM hoặc sheet dữ liệu trong file Excel.")

    workbook.close()

def template_report_type(filename: str, family: str) -> str:

    stem = Path(filename).stem

    if family == "rfi":

        match = re.search(r"\bRFI[\s-]*([A-Z0-9]+)", stem, flags=re.IGNORECASE)

    elif family == "rp":

        match = re.search(r"\bRP\s*-\s*([A-Za-z]+)(?![A-Za-z0-9])", stem, flags=re.IGNORECASE)

    else:

        raise UserError("Không nhận diện được nhóm biểu mẫu.")

    if match is None:

        expected = "RFI-LOẠI" if family == "rfi" else "RP-LOẠI"

        raise UserError(f"Tên file biểu mẫu phải có mã loại, ví dụ {expected}.xlsx.")

    return match.group(1).upper()

def validate_template_upload(filename: str, content: bytes, family: str) -> str:

    if len(content) > MAX_UPLOAD_BYTES:

        raise UserError("File biểu mẫu vượt giới hạn 100 MB.")

    if Path(filename).suffix.lower() not in {".xlsx", ".xlsm"}:

        raise UserError("Chỉ nhận file biểu mẫu .xlsx hoặc .xlsm.")

    report_type = template_report_type(filename, family)

    if not zipfile.is_zipfile(__import__("io").BytesIO(content)):

        raise UserError("File biểu mẫu không phải workbook Excel hợp lệ.")

    try:

        workbook = load_workbook(

            __import__("io").BytesIO(content),

            read_only=True,

            data_only=False,

            keep_vba=Path(filename).suffix.lower() == ".xlsm",

        )

    except Exception as exc:

        raise UserError(f"Không đọc được file biểu mẫu Excel: {exc}") from exc

    try:

        if not workbook.sheetnames:

            raise UserError("Workbook biểu mẫu không có sheet.")

        if family == "rfi":

            has_selector = any(

                "RFI-" in str(sheet["A6"].value or "").upper()

                for sheet in workbook.worksheets

            )

            if not has_selector and report_type not in {"MTR", "PPP", "MRIR", "DIC"}:

                raise UserError(

                    "Không nhận diện được ô chọn RFI của biểu mẫu (thường là ô A6)."

                )

    finally:

        workbook.close()

    return report_type

def available_templates(location: str) -> dict[str, list[dict[str, str]]]:

    result: dict[str, list[dict[str, str]]] = {"rfi": [], "rp": []}

    for family in ("rfi", "rp"):

        folder = site_template_dir(location, family)

        if not folder.is_dir():

            continue

        found: dict[str, tuple[float, dict[str, str]]] = {}

        for path in folder.iterdir():

            if path.suffix.lower() not in {".xlsx", ".xlsm"}:

                continue

            if family == "rfi" and path.name.startswith("FORM CHOOSE"):

                continue

            try:

                report_type = template_report_type(path.name, family)

            except UserError:

                continue

            mtime = path.stat().st_mtime

            if report_type not in found or mtime > found[report_type][0]:

                found[report_type] = (

                    mtime,

                    {

                        "type": report_type,

                        "filename": path.name,

                        "modified": datetime.fromtimestamp(mtime).isoformat(timespec="minutes"),

                    },

                )

        result[family] = [found[key][1] for key in sorted(found)]

    return result

def read_multipart(handler: BaseHTTPRequestHandler) -> tuple[dict[str, str], list[tuple[str, bytes]]]:

    content_length = int(handler.headers.get("Content-Length", "0"))

    if content_length < 1 or content_length > MAX_UPLOAD_BYTES + 4 * 1024 * 1024:

        raise UserError("Dữ liệu tải lên rỗng hoặc vượt giới hạn 100 MB.")

    content_type = handler.headers.get("Content-Type", "")

    if "multipart/form-data" not in content_type:

        raise UserError("Yêu cầu tải file không đúng định dạng.")

    raw = handler.rfile.read(content_length)

    message = BytesParser(policy=policy.default).parsebytes(

        b"Content-Type: " + content_type.encode("ascii") + b"\r\n"

        b"MIME-Version: 1.0\r\n\r\n" + raw

    )

    if not message.is_multipart():

        raise UserError("Không đọc được file tải lên.")

    fields: dict[str, str] = {}

    uploads: list[tuple[str, bytes]] = []

    for part in message.iter_parts():

        field_name = part.get_param("name", header="content-disposition")

        filename = part.get_filename()

        payload = part.get_payload(decode=True) or b""

        if filename:

            uploads.append((filename, payload))

        elif field_name:

            fields[field_name] = payload.decode("utf-8", errors="replace")

    return fields, uploads


# ==============================================================================
# AUTHENTICATION & ACCESS CONTROL (RBAC) WITH GMAIL APPROVAL SYSTEM
# ==============================================================================
USERS_FILE = DATA_DIR / "users.json"
DEFAULT_USERS = {
    "letruong28491@gmail.com": {
        "email": "letruong28491@gmail.com",
        "name": "Quản Trị Viên (Lê Trường)",
        "password": "Admin@2026",
        "status": "approved",
        "sites": ["*"],
        "role": "admin",
        "created_at": "2026-11-05 08:00:00"
    },
    "letruonggdct6nt@gmail.com": {
        "email": "letruonggdct6nt@gmail.com",
        "name": "Lê Trường (QC TTC11)",
        "password": "QC@2026",
        "status": "approved",
        "sites": ["TTC11"],
        "role": "qc",
        "created_at": "2026-11-05 08:00:00"
    },
    "minhnguyen1432000@gmail.com": {
        "email": "minhnguyen1432000@gmail.com",
        "name": "Nguyễn Văn Minh (QC TTC11)",
        "password": "QC@2026",
        "status": "approved",
        "sites": ["TTC11"],
        "role": "qc",
        "created_at": "2026-11-06 00:44:47"
    }
}

SESSIONS_FILE = DATA_DIR / "sessions.json"

def load_sessions() -> dict[str, dict[str, object]]:
    if SESSIONS_FILE.exists():
        try:
            with open(SESSIONS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
    return {}

def save_sessions(sessions: dict[str, dict[str, object]]) -> None:
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(SESSIONS_FILE, "w", encoding="utf-8") as f:
            json.dump(sessions, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[Sessions] Warning: Could not save sessions: {e}")

SESSIONS: dict[str, dict[str, object]] = load_sessions()

def load_users() -> dict[str, dict[str, object]]:
    if USERS_FILE.exists():
        try:
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict) and data:
                    # Remove legacy non-Gmail accounts
                    for old_k in ("ttc11", "ttc15", "ttc15a"):
                        data.pop(old_k, None)
                    modified = False
                    for k, v in data.items():
                        # Auto-approve so registered users don't get locked out
                        if v.get("status") in ("pending", None):
                            v["status"] = "approved"
                            modified = True
                        if not v.get("sites"):
                            v["sites"] = ["*"]
                            modified = True
                        if "role" not in v:
                            v["role"] = "guest"
                            modified = True
                    if modified:
                        save_users(data)
                    return data
        except Exception:
            pass
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(DEFAULT_USERS, f, ensure_ascii=False, indent=2)
    return DEFAULT_USERS

def save_users(users: dict[str, dict[str, object]]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

def find_user_by_identity(identity: str) -> tuple[str | None, dict[str, object] | None]:
    clean_id = identity.strip().lower()
    users = load_users()
    if clean_id in users:
        return clean_id, users[clean_id]
    for k, u in users.items():
        if str(u.get("email", "")).strip().lower() == clean_id:
            return k, u
    return None, None

def authenticate_user(identity: str, password: str) -> tuple[dict[str, object] | None, str]:
    key, u = find_user_by_identity(identity)
    if not u or u.get("password") != password:
        return None, "Tên đăng nhập / Gmail hoặc mật khẩu không chính xác."
    
    status = u.get("status", "approved")
    if status == "pending":
        u["status"] = "approved"
        if not u.get("sites"):
            u["sites"] = ["*"]
        users = load_users()
        if key in users:
            users[key]["status"] = "approved"
            users[key]["sites"] = u["sites"]
            save_users(users)
        status = "approved"
    elif status == "rejected":
        return None, "Tài khoản của bạn đã bị từ chối hoặc vô hiệu hóa bởi Quản trị viên."

    sites = u.get("sites", [])
    if not sites:
        sites = ["*"]

    return {
        "key": key,
        "email": u.get("email", key),
        "username": key,
        "name": u.get("name", key),
        "role": u.get("role", "qc"),
        "sites": sites,
        "status": status
    }, ""

def check_role_permission(user: dict[str, object] | None, allowed_roles: tuple[str, ...]) -> tuple[bool, str]:
    if not user:
        return False, "Vui lòng đăng nhập để tiếp tục."
    role = str(user.get("role", "guest")).lower()
    if role not in allowed_roles:
        role_names = {
            "admin": "Quản trị viên (Admin)",
            "qc": "Kỹ sư QC",
            "construction": "Thi công",
            "guest": "Khách (Chỉ xem)"
        }
        name = role_names.get(role, role)
        return False, f"Tài khoản vai trò '{name}' không có quyền thực hiện thao tác này."
    return True, ""

def check_site_permission(user: dict[str, object] | None, site: str) -> bool:
    if not user:
        return False
    allowed = user.get("sites", [])
    if "*" in allowed:
        return True
    site_clean = str(site).strip().upper()
    return any(str(s).strip().upper() == site_clean for s in allowed)



def save_source_upload_meta(location: str, key: str, original_filename: str, file_size: int, updated_at: str, updated_by: str) -> None:
    try:
        folder = site_data_dir(location)
        folder.mkdir(parents=True, exist_ok=True)
        cfg_path = folder / "site_config.json"
        data = {}
        if cfg_path.is_file():
            try:
                data = json.loads(cfg_path.read_text(encoding="utf-8"))
            except Exception:
                data = {}
        if "sources_meta" not in data or not isinstance(data["sources_meta"], dict):
            data["sources_meta"] = {}
        data["sources_meta"][key] = {
            "original_filename": original_filename,
            "size": file_size,
            "updated_at": updated_at,
            "updated_by": updated_by
        }
        cfg_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        print(f"[UploadMeta] Warning: Could not save upload meta: {e}")

def get_source_upload_meta(location: str, key: str) -> dict:
    try:
        folder = site_data_dir(location)
        cfg_path = folder / "site_config.json"
        if cfg_path.is_file():
            data = json.loads(cfg_path.read_text(encoding="utf-8"))
            meta = data.get("sources_meta", {})
            if isinstance(meta, dict):
                return meta.get(key, {})
    except Exception:
        pass
    return {}

def get_site_progress(location: str) -> dict[str, object]:
    path = source_path("controlog", location)
    if not path.is_file():
        for cand in [Path(f"Fome_Sites/{location}/Fome_Web_Data/Controlog WM {location}.xlsx"),
                     Path(f"Controlog WM {location}.xlsx"),
                     DATA_DIR / f"Controlog WM {location}.xlsx"]:
            if cand.is_file():
                path = cand
                break

    stages = []
    total_joints = 0
    completed_joints = 0
    active_joints = 0
    sample_joints = []

    if path.is_file():
        wb = None
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            if "Hoàn thành Joint" in wb.sheetnames:
                ws = wb["Hoàn thành Joint"]
                val_total = ws.cell(row=3, column=7).value
                val_comp = ws.cell(row=3, column=11).value
                val_active = ws.cell(row=3, column=15).value

                data_rows = list(ws.iter_rows(min_row=7, max_row=ws.max_row, values_only=True))
                valid_rows = [r for r in data_rows if len(r) >= 3 and r[2] is not None and str(r[2]).strip() != ""]
                total_joints = len(valid_rows)

                def count_step(date_idx, rep_idx):
                    cnt = 0
                    for r in valid_rows:
                        has_d = date_idx < len(r) and r[date_idx] is not None and str(r[date_idx]).strip() not in ("", "None", "0", "Chưa thực hiện")
                        has_r = rep_idx < len(r) and r[rep_idx] is not None and str(r[rep_idx]).strip() not in ("", "None", "0", "Chưa thực hiện")
                        if has_d or has_r:
                            cnt += 1
                    return cnt

                cnt_fitup = count_step(4, 5)
                cnt_weld = count_step(6, 7)
                cnt_visual = count_step(8, 9)
                cnt_ndt = count_step(10, 11)
                cnt_surface = count_step(12, 13)
                cnt_coating = count_step(14, 15)
                
                # Check lowering / completion column (Col 23, 24) or fall back to HDPUF (Col 19, 20)
                cnt_lowering = count_step(22, 23)
                cnt_complete = cnt_lowering if cnt_lowering > 0 else count_step(18, 19)

                stages = [
                    {"name": "Gá lắp (Fit-up)", "count": cnt_fitup, "icon": "🔧"},
                    {"name": "Hàn (Welding)", "count": cnt_weld, "icon": "🔥"},
                    {"name": "Ngoại quan (Visual)", "count": cnt_visual, "icon": "🔍"},
                    {"name": "NDT (AUT/RT)", "count": cnt_ndt, "icon": "📡"},
                    {"name": "Làm sạch (Surface)", "count": cnt_surface, "icon": "🧹"},
                    {"name": "Sơn phủ (Painting/Coating)", "count": cnt_coating, "icon": "🛡️"},
                    {"name": "Nghiệm thu (RFI/RP)", "count": cnt_complete, "icon": "📑"},
                ]

                try:
                    completed_joints = int(val_comp) if val_comp is not None else cnt_complete
                except Exception:
                    completed_joints = cnt_complete

                try:
                    active_joints = int(val_active) if val_active is not None else max(0, cnt_fitup - completed_joints)
                except Exception:
                    active_joints = max(0, cnt_fitup - completed_joints)

                for r in valid_rows:
                    j_no = r[1]
                    j_code = str(r[2])
                    kp = r[3]
                    kp_str = f"{float(kp):.1f}" if isinstance(kp, (int, float)) else str(kp or "")

                    fitup_str = str(r[5] or (str(r[4])[:10] if r[4] is not None else ""))
                    weld_str = str(r[7] or (str(r[6])[:10] if r[6] is not None else ""))
                    visual_str = str(r[9] or (str(r[8])[:10] if r[8] is not None else ""))
                    ndt_str = str(r[11] or (str(r[10])[:10] if r[10] is not None else ""))
                    surface_str = str(r[13] or (str(r[12])[:10] if r[12] is not None else ""))
                    coating_str = str(r[15] or (str(r[14])[:10] if r[14] is not None else ""))

                    sample_joints.append({
                        "no": j_no,
                        "code": j_code,
                        "kp": kp_str,
                        "fitup": fitup_str if fitup_str not in ("None", "0", "") else "",
                        "weld": weld_str if weld_str not in ("None", "0", "") else "",
                        "visual": visual_str if visual_str not in ("None", "0", "") else "",
                        "ndt": ndt_str if ndt_str not in ("None", "0", "") else "",
                        "surface": surface_str if surface_str not in ("None", "0", "") else "",
                        "coating": coating_str if coating_str not in ("None", "0", "") else ""
                    })

            elif "Check" in wb.sheetnames:
                ws = wb["Check"]
                rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True))
                valid_rows = [r for r in rows if len(r) >= 2 and r[1] is not None and str(r[1]).strip() != ""]
                total_joints = len(valid_rows)

                def count_c(c_idx):
                    return sum(1 for r in valid_rows if c_idx < len(r) and r[c_idx] is not None and str(r[c_idx]).strip() not in ("", "None", "0", "False"))

                cnt_fitup = count_c(2)
                cnt_weld = count_c(3)
                cnt_visual = count_c(4)
                cnt_ndt = count_c(5)
                cnt_surface = count_c(6)
                cnt_coating = count_c(7)
                cnt_complete = count_c(11)

                stages = [
                    {"name": "Gá lắp (Fit-up)", "count": cnt_fitup, "icon": "🔧"},
                    {"name": "Hàn (Welding)", "count": cnt_weld, "icon": "🔥"},
                    {"name": "Ngoại quan (Visual)", "count": cnt_visual, "icon": "🔍"},
                    {"name": "NDT (AUT/RT)", "count": cnt_ndt, "icon": "📡"},
                    {"name": "Làm sạch (Surface)", "count": cnt_surface, "icon": "🧹"},
                    {"name": "Sơn phủ (Painting/Coating)", "count": cnt_coating, "icon": "🛡️"},
                    {"name": "Nghiệm thu (RFI/RP)", "count": cnt_complete, "icon": "📑"},
                ]
                completed_joints = sum(1 for r in valid_rows if len(r) >= 13 and (r[12] is True or str(r[12]).lower() == "true"))
                if completed_joints == 0:
                    completed_joints = cnt_complete
                active_joints = max(0, cnt_fitup - completed_joints)

                for r in valid_rows:
                    sample_joints.append({
                        "no": r[0],
                        "code": str(r[1]),
                        "kp": "",
                        "fitup": str(r[2])[:10] if r[2] is not None else "",
                        "weld": str(r[3])[:10] if r[3] is not None else "",
                        "visual": str(r[4])[:10] if r[4] is not None else "",
                        "ndt": str(r[5])[:10] if r[5] is not None else "",
                        "surface": str(r[6])[:10] if r[6] is not None else "",
                        "coating": str(r[7])[:10] if r[7] is not None else ""
                    })
        except Exception:
            pass
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass

    try:
        recs, _, _ = load_records(location)
    except Exception:
        recs = []

    return {
        "location": location,
        "total_joints": total_joints,
        "completed_joints": completed_joints,
        "active_joints": active_joints,
        "total_rfis": len(recs),
        "stages": stages,
        "sample_joints": sample_joints
    }


def get_site_data_version_info(location: str, records: list[dict] | None = None) -> dict[str, object]:
    folder = site_data_dir(location)
    cfg_path = folder / "site_config.json"
    site_meta = {}
    if cfg_path.is_file():
        try:
            site_meta = json.loads(cfg_path.read_text(encoding="utf-8"))
        except Exception:
            pass

    sources_meta = site_meta.get("sources_meta", {}) if isinstance(site_meta, dict) else {}
    
    latest_dt = None
    latest_time_str = ""
    latest_updater = ""
    latest_file = ""

    for k, v in sources_meta.items():
        if isinstance(v, dict):
            u_at = str(v.get("updated_at", "")).strip()
            u_by = str(v.get("updated_by", "")).strip()
            f_name = str(v.get("original_filename", "")).strip()
            if u_at:
                for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M"):
                    try:
                        dt = datetime.strptime(u_at, fmt)
                        if latest_dt is None or dt > latest_dt:
                            latest_dt = dt
                            latest_time_str = u_at
                            latest_updater = u_by
                            latest_file = f_name
                        break
                    except Exception:
                        pass

    for key in SOURCES:
        p = source_path(key, location)
        if p.is_file():
            st = p.stat()
            dt = datetime.fromtimestamp(st.st_mtime)
            if latest_dt is None or dt > latest_dt:
                latest_dt = dt
                latest_time_str = dt.strftime("%d/%m/%Y %H:%M:%S")
                if not latest_file:
                    latest_file = p.name

    earliest_rec_date = ""
    latest_rec_date = ""
    if records:
        valid_dates = sorted([r["date"] for r in records if r.get("date")])
        if valid_dates:
            try:
                earliest_rec_date = datetime.strptime(valid_dates[0], "%Y-%m-%d").strftime("%d/%m/%Y")
                latest_rec_date = datetime.strptime(valid_dates[-1], "%Y-%m-%d").strftime("%d/%m/%Y")
            except Exception:
                earliest_rec_date = valid_dates[0]
                latest_rec_date = valid_dates[-1]

    if latest_dt:
        cutoff_date = latest_dt.strftime("%d/%m/%Y")
    elif latest_rec_date:
        cutoff_date = latest_rec_date
    else:
        cutoff_date = datetime.now().strftime("%d/%m/%Y")

    if not latest_time_str and latest_dt:
        latest_time_str = latest_dt.strftime("%d/%m/%Y %H:%M:%S")

    date_range_str = f"{earliest_rec_date} → {latest_rec_date}" if earliest_rec_date and latest_rec_date else (latest_rec_date or cutoff_date)

    summary_msg = f"Dữ liệu đang nhận thông tin ở ngày {cutoff_date}"
    if latest_time_str:
        summary_msg += f" (Cập nhật lúc {latest_time_str}"
        if latest_updater:
            summary_msg += f" bởi {latest_updater}"
        summary_msg += ")"

    return {
        "cutoff_date": cutoff_date,
        "last_update_time": latest_time_str,
        "last_updater": latest_updater or "Kỹ sư",
        "last_file": latest_file or "Controlog WM.xlsx",
        "total_records": len(records) if records else 0,
        "earliest_date": earliest_rec_date,
        "latest_date": latest_rec_date,
        "date_range": date_range_str,
        "summary_message": summary_msg
    }


class FomeHandler(BaseHTTPRequestHandler):

    server_version = "FomeLocal/1.0"

    def get_current_user(self) -> dict[str, object] | None:
        cookie_header = self.headers.get("Cookie", "")
        token = None
        if "fome_session=" in cookie_header:
            for part in cookie_header.split(";"):
                part = part.strip()
                if part.startswith("fome_session="):
                    token = part.split("=", 1)[1]
                    break
        if not token:
            auth_header = self.headers.get("Authorization", "")
            if auth_header.startswith("Bearer "):
                token = auth_header.split(" ", 1)[1].strip()
        if token and token in SESSIONS:
            return SESSIONS[token]
        return None

    def handle_login(self) -> None:
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length).decode("utf-8")
            data = json.loads(body) if body else {}
            username = str(data.get("username", "")).strip()
            password = str(data.get("password", "")).strip()
            user, err_msg = authenticate_user(username, password)
            if not user:
                status_code = 403 if "từ chối" in err_msg else 401
                self.send_json(status_code, {"error": err_msg})
                return
            token = hashlib.sha256(f"{user['username']}:{datetime.now().isoformat()}:{random.random()}".encode("utf-8")).hexdigest()
            SESSIONS[token] = user
            save_sessions(SESSIONS)
            cookie_val = f"fome_session={token}; Path=/; HttpOnly; SameSite=Lax; Max-Age=2592000"
            payload = json.dumps({"success": True, "user": user, "token": token}, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(payload)))
            self.send_header("Set-Cookie", cookie_val)
            self.send_header("Cache-Control", "no-store")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(payload)
        except Exception as exc:
            self.send_json(500, {"error": f"Lỗi đăng nhập: {exc}"})

    def handle_register(self) -> None:
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length).decode("utf-8")
            data = json.loads(body) if body else {}
            email = str(data.get("email", "")).strip().lower()
            name = str(data.get("name", "")).strip()
            password = str(data.get("password", "")).strip()
            requested_sites = data.get("requested_sites", [])

            if not email or "@" not in email:
                self.send_json(400, {"error": "Vui lòng nhập địa chỉ Gmail hợp lệ."})
                return
            if not name:
                self.send_json(400, {"error": "Vui lòng nhập họ và tên của bạn."})
                return
            if len(password) < 6:
                self.send_json(400, {"error": "Mật khẩu phải có tối thiểu 6 ký tự."})
                return

            users = load_users()
            if email in users or any(u.get("email", "").lower() == email for u in users.values()):
                self.send_json(400, {"error": f"Địa chỉ Gmail '{email}' đã tồn tại trên hệ thống. Bạn có thể bấm Đăng nhập ngay."})
                return

            requested_role = str(data.get("role", "qc")).strip().lower()
            if requested_role not in ("admin", "qc", "construction", "guest"):
                requested_role = "qc"
            
            allocated_sites = requested_sites if (isinstance(requested_sites, list) and requested_sites) else ["*"]
            new_user = {
                "email": email,
                "name": name,
                "password": password,
                "status": "approved",
                "role": requested_role,
                "sites": allocated_sites,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            users[email] = new_user
            save_users(users)

            user_payload = {
                "key": email,
                "email": email,
                "username": email,
                "name": name,
                "role": requested_role,
                "sites": allocated_sites,
                "status": "approved"
            }
            token = hashlib.sha256(f"{email}:{datetime.now().isoformat()}:{random.random()}".encode("utf-8")).hexdigest()
            SESSIONS[token] = user_payload
            save_sessions(SESSIONS)

            cookie_val = f"fome_session={token}; Path=/; HttpOnly; SameSite=Lax; Max-Age=2592000"
            payload = json.dumps({
                "success": True,
                "user": user_payload,
                "token": token,
                "message": f"Đăng ký tài khoản thành công! Bạn đã được tự động đăng nhập với vai trò '{requested_role.upper()}'."
            }, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(payload)))
            self.send_header("Set-Cookie", cookie_val)
            self.send_header("Cache-Control", "no-store")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(payload)
        except Exception as exc:
            self.send_json(500, {"error": f"Lỗi đăng ký: {exc}"})

    def handle_admin_get_users(self) -> None:
        user = self.get_current_user()
        if not user or user.get("role") != "admin":
            self.send_json(403, {"error": "Chỉ Quản trị viên (Admin) mới có quyền truy cập."})
            return
        users = load_users()
        user_list = []
        for key, val in users.items():
            user_list.append({
                "id": key,
                "email": val.get("email", key),
                "name": val.get("name", key),
                "status": val.get("status", "approved"),
                "role": val.get("role", "user"),
                "sites": val.get("sites", []),
                "created_at": val.get("created_at", "")
            })
        self.send_json(200, {"users": user_list, "available_sites": list_sites()})

    def handle_admin_update_user(self) -> None:
        user = self.get_current_user()
        if not user or user.get("role") != "admin":
            self.send_json(403, {"error": "Chỉ Quản trị viên (Admin) mới có quyền truy cập."})
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length).decode("utf-8")
            data = json.loads(body) if body else {}
            target_id = str(data.get("id", "")).strip().lower()
            new_status = data.get("status")
            new_sites = data.get("sites")
            new_role = data.get("role")

            users = load_users()
            if target_id not in users:
                self.send_json(404, {"error": "Không tìm thấy người dùng."})
                return

            if target_id == "letruong28491@gmail.com":
                users[target_id]["status"] = "approved"
                users[target_id]["role"] = "admin"
                users[target_id]["sites"] = ["*"]
            else:
                if new_status:
                    users[target_id]["status"] = str(new_status)
                if new_sites is not None and isinstance(new_sites, list):
                    users[target_id]["sites"] = new_sites
                if new_role:
                    clean_role = str(new_role).strip().lower()
                    if clean_role in ("admin", "qc", "construction", "guest"):
                        users[target_id]["role"] = clean_role
                        if clean_role == "admin" and not users[target_id].get("sites"):
                            users[target_id]["sites"] = ["*"]

            save_users(users)

            # Update live sessions if active
            for tok, s_user in list(SESSIONS.items()):
                if s_user.get("username") == target_id or s_user.get("email") == target_id:
                    s_user["status"] = users[target_id].get("status", "approved")
                    s_user["sites"] = users[target_id].get("sites", [])
                    s_user["role"] = users[target_id].get("role", "guest")

            self.send_json(200, {
                "success": True,
                "message": f"Đã cập nhật phê duyệt và phân quyền cho {users[target_id].get('name')} ({target_id}) thành công."
            })
        except Exception as exc:
            self.send_json(500, {"error": f"Lỗi cập nhật người dùng: {exc}"})

    def handle_admin_delete_user(self) -> None:
        user = self.get_current_user()
        if not user or user.get("role") != "admin":
            self.send_json(403, {"error": "Chỉ Quản trị viên (Admin) mới có quyền truy cập."})
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length).decode("utf-8")
            data = json.loads(body) if body else {}
            target_id = str(data.get("id", "")).strip().lower()
            if target_id in ("letruong28491@gmail.com", "admin"):
                self.send_json(400, {"error": "Không thể xóa tài khoản Quản trị viên mặc định."})
                return
            users = load_users()
            if target_id in users:
                del users[target_id]
                save_users(users)
            # Remove any active sessions
            for tok, s_user in list(SESSIONS.items()):
                if s_user.get("username") == target_id or s_user.get("email") == target_id:
                    del SESSIONS[tok]
            self.send_json(200, {"success": True, "message": f"Đã xóa tài khoản {target_id} khỏi hệ thống."})
        except Exception as exc:
            self.send_json(500, {"error": f"Lỗi xóa người dùng: {exc}"})

    def handle_logout(self) -> None:
        auth_header = self.headers.get("Authorization", "")
        if auth_header.startswith("Bearer "):
            tok = auth_header.split(" ", 1)[1].strip()
            SESSIONS.pop(tok, None)
        cookie_header = self.headers.get("Cookie", "")
        for part in cookie_header.split(";"):
            part = part.strip()
            if part.startswith("fome_session="):
                tok = part.split("=", 1)[1]
                SESSIONS.pop(tok, None)
        save_sessions(SESSIONS)
        payload = b'{"success": true}'
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Set-Cookie", "fome_session=; Path=/; Expires=Thu, 01 Jan 1970 00:00:00 GMT")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(payload)


    def send_json(self, status: int, payload: dict[str, object]) -> None:

        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")

        self.send_response(status)

        self.send_header("Content-Type", "application/json; charset=utf-8")

        self.send_header("Content-Length", str(len(data)))

        self.send_header("Cache-Control", "no-store")

        self.send_header("Access-Control-Allow-Origin", "*")

        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")

        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")

        self.end_headers()

        self.wfile.write(data)

    def do_OPTIONS(self) -> None:

        self.send_response(204)

        self.send_header("Access-Control-Allow-Origin", "*")

        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")

        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")

        self.end_headers()

    def verify_local_request(self) -> bool:

        return True

    def do_GET(self) -> None:
        parsed = urlparse(self.path)

        if parsed.path in ("/", "", "/index.html"):
            try:
                page = (WEB_DIR / "index.html").read_bytes()
            except OSError:
                self.send_response(500)
                self.send_header("Content-Type", "text/plain; charset=utf-8")
                self.end_headers()
                self.wfile.write("Không tìm thấy giao diện web.".encode("utf-8"))
                return
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(page)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(page)
            return

        if parsed.path == "/favicon.ico":
            self.send_response(204)
            self.end_headers()
            return

        if parsed.path == "/api/me":
            user = self.get_current_user()
            if user:
                # Count pending users if admin
                pending_count = 0
                if user.get("role") == "admin":
                    all_u = load_users()
                    pending_count = sum(1 for u in all_u.values() if u.get("status") == "pending")
                self.send_json(200, {"authenticated": True, "user": user, "pending_count": pending_count})
            else:
                self.send_json(200, {"authenticated": False})
            return

        if parsed.path == "/api/admin/users":
            self.handle_admin_get_users()
            return

        if parsed.path == "/api/export-site-sources":
            self.handle_export_site_sources()
            return

        if parsed.path == "/api/sites":
            user = self.get_current_user()
            if not user:
                self.send_json(401, {"error": "Vui lòng đăng nhập để tiếp tục.", "authenticated": False})
                return
            all_sites = list_sites()
            if "*" in user.get("sites", []):
                allowed_sites = all_sites
            else:
                allowed_sites = [s for s in all_sites if s in user.get("sites", [])]
                for s in user.get("sites", []):
                    if s not in allowed_sites:
                        allowed_sites.append(s)
            default_site = allowed_sites[0] if allowed_sites else DEFAULT_SITE
            self.send_json(200, {
                "sites": allowed_sites,
                "default": default_site,
                "user": user
            })
            return

            try:

                self.send_json(

                    200,

                    {

                        "sites": list_sites(),

                        "default": get_active_site(),

                    },

                )

            except Exception as exc:

                self.send_json(400, {"error": str(exc)})

            return

        if parsed.path == "/api/active-site":

            try:

                self.send_json(200, {"location": get_active_site()})

            except Exception as exc:

                self.send_json(400, {"error": str(exc)})

            return

        if parsed.path == "/api/site-drawing":
            user = self.get_current_user()
            if not user:
                self.send_json(401, {"error": "Vui lòng đăng nhập.", "authenticated": False})
                return
            try:
                query = parse_qs(parsed.query)
                req_loc = query.get("location", [""])[0] or get_active_site()
                info = get_site_drawing_info(req_loc)
                self.send_json(200, info)
            except Exception as exc:
                self.send_json(500, {"error": f"Lỗi đọc file bản vẽ: {exc}"})
            return

        if parsed.path == "/api/nktc-days":
            user = self.get_current_user()
            if not user:
                self.send_json(401, {"error": "Vui lòng đăng nhập.", "authenticated": False})
                return
            try:
                query = parse_qs(parsed.query)
                req_loc = query.get("location", [""])[0] or get_active_site()
                if not check_site_permission(user, req_loc):
                    self.send_json(403, {"error": f"Tài khoản không có quyền truy cập vị trí {req_loc}."})
                    return
                if not HAS_NKTC:
                    self.send_json(500, {"error": "Mô-đun Nhật ký thi công chưa được nạp."})
                    return
                tpl_info = nktc_processor.get_site_nktc_template_info(req_loc)
                _, days_list, sorted_months = nktc_processor.get_all_site_days(req_loc)
                self.send_json(200, {
                    "location": req_loc,
                    "template_info": tpl_info,
                    "days": days_list,
                    "months": [{"value": m[0], "label": f"Tháng {m[1]}"} for m in sorted_months]
                })
            except Exception as exc:
                self.send_json(500, {"error": f"Lỗi nạp danh sách ngày thi công: {exc}"})
            return

        if parsed.path == "/api/progress":
            user = self.get_current_user()
            if not user:
                self.send_json(401, {"error": "Vui lòng đăng nhập.", "authenticated": False})
                return
            try:
                query = parse_qs(parsed.query)
                req_loc = query.get("location", [""])[0] or get_active_site()
                if not check_site_permission(user, req_loc):
                    self.send_json(403, {"error": f"Tài khoản không có quyền truy cập vị trí {req_loc}."})
                    return
                data = get_site_progress(req_loc)
                self.send_json(200, data)
            except Exception as exc:
                self.send_json(500, {"error": f"Lỗi đọc dữ liệu tiến độ: {exc}"})
            return

        if parsed.path == "/api/status":
            user = self.get_current_user()
            if not user:
                self.send_json(401, {"error": "Vui lòng đăng nhập.", "authenticated": False})
                return
            try:
                query = parse_qs(parsed.query)
                req_loc = query.get("location", [""])[0] or get_active_site()
                if not check_site_permission(user, req_loc):
                    self.send_json(403, {"error": f"Tài khoản {user.get('name')} không có quyền truy cập vị trí {req_loc}."})
                    return

                requested_location = query.get("location", [""])[0]

                if not requested_location:

                    location = get_active_site()

                else:

                    location = validate_site_code(requested_location, require_existing=False)

                    set_active_site(location)

                warning = ""

                records = []

                try:

                    records, _, dup_ids = load_records(location)

                    if dup_ids:

                        warning = f"Số RFI bị trùng (đã bỏ qua dòng trùng): {', '.join(dup_ids)}"

                except Exception as exc:

                    warning = str(exc)

                sources = []
                for key in SOURCES:
                    path = source_path(key, location)
                    fname = path.name
                    stat = path.stat() if path.is_file() else None
                    meta = get_source_upload_meta(location, key)
                    sources.append(
                        {
                            "key": key,
                            "filename": fname,
                            "original_filename": meta.get("original_filename", fname),
                            "updated_by": meta.get("updated_by", ""),
                            "updated_at": meta.get("updated_at", ""),
                            "exists": stat is not None,
                            "size": stat.st_size if stat else 0,
                            "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%d/%m/%Y %H:%M") if stat else None,
                        }
                    )

                try:

                    site_config = get_site_config(location)

                except Exception:

                    site_config = {

                        "drawing_no": DEFAULT_DRAWING_NO,

                        "weld_map_no": DEFAULT_WELD_MAP_NO,

                    }

                try:

                    cached_tables = get_cached_rp_source_tables(location)

                except Exception:

                    cached_tables = {}

                if cached_tables and records:

                    for rec in records:

                        try:

                            d_no, wm_no = resolve_drawings_for_rfi(

                                rec["id"],

                                rec["type"],

                                cached_tables,

                                location,

                                fallback_drawing=site_config.get("drawing_no", DEFAULT_DRAWING_NO),

                                fallback_weld_map=site_config.get("weld_map_no", DEFAULT_WELD_MAP_NO),

                            )

                            rec["drawing_no"] = d_no

                            rec["weld_map_no"] = wm_no

                        except Exception:

                            rec["drawing_no"] = site_config.get("drawing_no", DEFAULT_DRAWING_NO)

                            rec["weld_map_no"] = site_config.get("weld_map_no", DEFAULT_WELD_MAP_NO)

                self.send_json(

                    200,

                    {

                        "sources": sources,

                        "records": records,

                        "count": len(records),

                        "location": location,

                        "station_code": station_code_for_location(location),

                        "drawing_no": site_config.get("drawing_no", DEFAULT_DRAWING_NO),

                        "weld_map_no": site_config.get("weld_map_no", DEFAULT_WELD_MAP_NO),

                        "engineer_name": site_config.get(

                            "engineer_name", default_engineer_for_location(location)

                        ),

                        "templates": available_templates(location),

                        "drawing_info": get_site_drawing_info(location),
                        "nktc_template_info": nktc_processor.get_site_nktc_template_info(location) if HAS_NKTC else None,
                        "data_version": get_site_data_version_info(location, records),

                        "warning": warning,

                    },

                )

            except Exception as exc:

                self.send_json(500, {"error": f"Lỗi đọc trạng thái: {exc}"})

            return

        self.send_error(404)

    def do_POST(self) -> None:

        if not self.verify_local_request():

            return

        if self.path == "/api/login":
            self.handle_login()
            return

        if self.path == "/api/register":
            self.handle_register()
            return

        if self.path == "/api/logout":
            self.handle_logout()
            return

        if self.path == "/api/admin/update-user":
            self.handle_admin_update_user()
            return

        if self.path == "/api/admin/delete-user":
            self.handle_admin_delete_user()
            return

        user = self.get_current_user()
        if not user:
            self.send_json(401, {"error": "Vui lòng đăng nhập để thực hiện thao tác.", "authenticated": False})
            return

        if self.path == "/api/site" and user.get("role") != "admin":
            self.send_json(403, {"error": "Chỉ tài khoản Quản trị viên (Admin) mới có quyền tạo vị trí mới."})
            return

        if self.path == "/api/upload":

            self.handle_upload()

            return

        if self.path == "/api/site":

            self.handle_create_site()

            return

        if self.path == "/api/active-site":

            self.handle_set_active_site()

            return

        if self.path == "/api/template":

            self.handle_template_upload()

            return

        if self.path == "/api/upload-nktc-template":
            self.handle_upload_nktc_template()
            return

        if self.path == "/api/export-nktc":
            self.handle_export_nktc()
            return

        if self.path == "/api/upload-drawing":

            self.handle_upload_drawing()

            return

        if self.path == "/api/export":

            self.handle_export()

            return

        if self.path == "/api/fix-station-code":

            self.handle_fix_station_code()

            return

        if self.path == "/api/site-config":

            self.handle_save_site_config()

            return

        self.send_error(404)

    def handle_upload_nktc_template(self) -> None:
        user = self.get_current_user()
        ok, err = check_role_permission(user, ("admin", "qc"))
        if not ok:
            self.send_json(403, {"error": err})
            return
        try:
            fields, uploads = read_multipart(self)
            location = validate_site_code(fields.get("location", DEFAULT_SITE), True)
            if not check_site_permission(user, location):
                self.send_json(403, {"error": f"Tài khoản không có quyền nạp biểu mẫu cho trạm {location}."})
                return
            if not uploads:
                raise UserError("Vui lòng chọn file mẫu Word (.docx) để tải lên.")
            filename, content = uploads[0]
            info = nktc_processor.save_site_nktc_template(location, filename, content)
            self.send_json(200, {
                "message": f"Đã nạp file mẫu Nhật ký thi công {info['filename']} cho trạm {location} thành công!",
                "template_info": info
            })
        except UserError as exc:
            self.send_json(400, {"error": str(exc)})
        except Exception as exc:
            self.send_json(500, {"error": f"Không thể lưu file mẫu: {exc}"})

    def handle_export_nktc(self) -> None:
        user = self.get_current_user()
        if not user:
            self.send_json(401, {"error": "Vui lòng đăng nhập để xuất báo cáo."})
            return
        u_role = str(user.get("role", "guest")).lower()
        if u_role == "guest":
            self.send_json(403, {"error": "Tài khoản quyền 'Khách' chỉ có quyền xem dữ liệu và tiến độ, không có quyền xuất hoặc tải báo cáo."})
            return
        ok, err = check_role_permission(user, ("admin", "qc", "construction"))
        if not ok:
            self.send_json(403, {"error": err})
            return
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            if content_length < 1 or content_length > 2 * 1024 * 1024:
                raise UserError("Dữ liệu yêu cầu không hợp lệ.")
            payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
            location = validate_site_code(payload.get("location", DEFAULT_SITE), True)
            if not check_site_permission(user, location):
                self.send_json(403, {"error": f"Tài khoản không có quyền xuất báo cáo cho trạm {location}."})
                return
            selected_dates = payload.get("dates", [])
            if not isinstance(selected_dates, list):
                selected_dates = []
            path, download_name = nktc_processor.export_nktc_document(location, selected_dates)
            content_type = (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                if path.suffix.lower() == ".docx"
                else "application/zip"
            )
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header(
                "Content-Disposition",
                f'attachment; filename="{download_name}"; filename*=UTF-8\'\'{quote(download_name)}',
            )
            self.send_header("Content-Length", str(path.stat().st_size))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            with path.open("rb") as stream:
                shutil.copyfileobj(stream, self.wfile)
        except UserError as exc:
            self.send_json(400, {"error": str(exc)})
        except Exception as exc:
            self.send_json(500, {"error": f"Lỗi tạo Nhật ký thi công: {exc}"})

    def handle_upload_drawing(self) -> None:
        user = self.get_current_user()
        ok, err = check_role_permission(user, ("admin", "qc"))
        if not ok:
            self.send_json(403, {"error": err})
            return
        try:
            fields, uploads = read_multipart(self)
            location = validate_site_code(fields.get("location", DEFAULT_SITE), True)
            if not check_site_permission(user, location):
                self.send_json(403, {"error": f"Tài khoản không có quyền nạp bản vẽ cho trạm {location}."})
                return
            if not uploads:
                raise UserError("Vui lòng chọn file bản vẽ PDF để tải lên.")
            filename, content = uploads[0]
            info = save_uploaded_drawing(location, filename, content)
            self.send_json(200, {
                "message": f"Đã nạp file bản vẽ {info['filename']} ({info['pages']} trang) cho trạm {location} thành công!",
                "drawing_info": info
            })
        except UserError as exc:
            self.send_json(400, {"error": str(exc)})
        except Exception as exc:
            self.send_json(500, {"error": f"Không thể lưu file bản vẽ: {exc}"})

    def handle_set_active_site(self) -> None:

        try:

            content_length = int(self.headers.get("Content-Length", "0"))

            payload = json.loads(self.rfile.read(content_length).decode("utf-8")) if content_length > 0 else {}

            location = validate_site_code(str(payload.get("location", "")), require_existing=False)

            set_active_site(location)

            self.send_json(200, {"location": location, "message": f"Đã lưu vị trí mặc định: {location}"})

        except Exception as exc:

            self.send_json(400, {"error": str(exc)})


    def handle_export_site_sources(self) -> None:
        user = self.get_current_user()
        if not user:
            self.send_json(401, {"error": "Vui lòng đăng nhập để tải dữ liệu."})
            return
        query = parse_qs(urlparse(self.path).query)
        loc_param = query.get("location", [""])[0] or get_active_site()
        location = validate_site_code(loc_param, True)
        if not check_site_permission(user, location):
            self.send_json(403, {"error": f"Tài khoản không có quyền truy cập dữ liệu của trạm {location}."})
            return
        folder = site_data_dir(location)
        if not folder.is_dir():
            self.send_json(404, {"error": f"Chưa có thư mục dữ liệu cho trạm {location}."})
            return
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for item in folder.iterdir():
                if item.is_file() and not item.name.startswith("~") and not item.name.endswith(".upload"):
                    zf.write(item, item.name)
            dwg_info = get_site_drawing_info(location)
            if dwg_info.get("has_drawing") and dwg_info.get("path"):
                dwg_p = Path(dwg_info["path"])
                if dwg_p.is_file():
                    zf.write(dwg_p, dwg_p.name)
        zip_bytes = buf.getvalue()
        zip_filename = f"Du_Lieu_Nguon_{location}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        self.send_response(200)
        self.send_header("Content-Type", "application/zip")
        self.send_header(
            "Content-Disposition",
            f'attachment; filename="{zip_filename}"; filename*=UTF-8\'\'{quote(zip_filename)}',
        )
        self.send_header("Content-Length", str(len(zip_bytes)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(zip_bytes)

    def handle_upload(self) -> None:
        user = self.get_current_user()
        ok, err = check_role_permission(user, ("admin", "qc"))
        if not ok:
            self.send_json(403, {"error": err})
            return
        try:

            fields, uploads = read_multipart(self)

            location = validate_site_code(fields.get("location", DEFAULT_SITE), True)
            if not check_site_permission(user, location):
                self.send_json(403, {"error": f"Tài khoản không có quyền tải file lên trạm {location}."})
                return

            key = fields.get("source", "")

            if key not in SOURCES:

                raise UserError("Hãy chọn một trong hai file dữ liệu nguồn.")

            if len(uploads) != 1:

                raise UserError("Mỗi lần chỉ tải lên một file nguồn.")

            filename, content = uploads[0]

            expected = source_filename(key, location)

            if Path(filename.replace("\\", "/")).suffix.lower() != ".xlsx":

                raise UserError("File dữ liệu nguồn phải có định dạng .xlsx.")

            validate_uploaded_source(key, content, location)

            destination = source_path(key, location)

            destination.parent.mkdir(parents=True, exist_ok=True)

            if destination.exists():

                backup_current_source(key, location)

            temporary = destination.with_name(destination.name + ".upload")
            temporary.write_bytes(content)

            # Retry loop for file replacement on Windows
            gc.collect()
            replaced = False
            last_err = None
            for _ in range(5):
                try:
                    temporary.replace(destination)
                    replaced = True
                    break
                except PermissionError as p_err:
                    last_err = p_err
                    gc.collect()
                    time.sleep(0.15)
                except Exception as ex:
                    last_err = ex
                    break
            if not replaced:
                try:
                    destination.write_bytes(content)
                    if temporary.exists():
                        temporary.unlink(missing_ok=True)
                except Exception:
                    raise UserError(f"Không thể ghi đè file trên máy chủ do file đang được mở hoặc tạm khóa: {last_err}")

            try:
                destination.touch()
            except Exception:
                pass

            # Lưu lại siêu dữ liệu bản cập nhật cuối cùng (người tải, thời gian, tên file gốc)
            now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            u_name = str(user.get("name") or user.get("email") or "Kỹ sư")
            save_source_upload_meta(location, key, filename, len(content), now_str, u_name)

            _RP_SOURCE_TABLES_CACHE.pop(location, None)
            set_active_site(location)

            self.send_json(
                200,
                {
                    "message": f"Đã cập nhật và đồng bộ file {destination.name} cho {location} (Cập nhật bởi: {u_name}).",
                    "filename": destination.name,
                    "location": location,
                    "updated_at": now_str,
                    "updated_by": u_name,
                },
            )

        except UserError as exc:

            self.send_json(400, {"error": str(exc)})

        except Exception as exc:

            self.send_json(500, {"error": f"Không cập nhật được file nguồn: {exc}"})

    def handle_create_site(self) -> None:

        try:

            content_length = int(self.headers.get("Content-Length", "0"))

            if content_length < 1 or content_length > 16 * 1024:

                raise UserError("Thông tin vị trí không hợp lệ.")

            payload = json.loads(self.rfile.read(content_length).decode("utf-8"))

            location_code = str(payload.get("location", ""))

            drawing_no = str(payload.get("drawing_no", "")).strip()

            weld_map_no = str(payload.get("weld_map_no", "")).strip()

            engineer_name = str(payload.get("engineer_name", "")).strip()

            code = create_site(location_code, drawing_no, weld_map_no, engineer_name)

            set_active_site(code)

            self.send_json(201, {"location": code, "message": f"Đã tạo hồ sơ {code}."})

        except UserError as exc:

            self.send_json(400, {"error": str(exc)})

        except Exception as exc:

            self.send_json(500, {"error": f"Không tạo được hồ sơ vị trí: {exc}"})

    def handle_save_site_config(self) -> None:
        user = self.get_current_user()
        ok, err = check_role_permission(user, ("admin", "qc"))
        if not ok:
            self.send_json(403, {"error": err})
            return
        try:

            content_length = int(self.headers.get("Content-Length", "0"))

            if content_length < 1 or content_length > 64 * 1024:

                raise UserError("Dữ liệu cấu hình không hợp lệ.")

            payload = json.loads(self.rfile.read(content_length).decode("utf-8"))

            location = validate_site_code(payload.get("location", DEFAULT_SITE), True)

            drawing_no = str(payload.get("drawing_no", "")).strip()

            weld_map_no = str(payload.get("weld_map_no", "")).strip()

            engineer_name = str(payload.get("engineer_name", "")).strip() if "engineer_name" in payload else None

            res = update_all_rp_drawings_for_site(location, drawing_no, weld_map_no, engineer_name)

            set_active_site(location)

            self.send_json(

                200,

                {

                    "message": f"Đã cập nhật số bản vẽ & kỹ sư phụ trách ({res.get('engineer_name', '')}) cho vị trí {location}.",

                    "config": {

                        "drawing_no": res["drawing_no"],

                        "weld_map_no": res["weld_map_no"],

                        "engineer_name": res.get("engineer_name", ""),

                    },

                    "result": res,

                },

            )

        except UserError as exc:

            self.send_json(400, {"error": str(exc)})

        except Exception as exc:

            self.send_json(500, {"error": f"Không lưu được cấu hình: {exc}"})

    def handle_template_upload(self) -> None:
        user = self.get_current_user()
        ok, err = check_role_permission(user, ("admin", "qc"))
        if not ok:
            self.send_json(403, {"error": err})
            return
        try:

            fields, uploads = read_multipart(self)

            location = validate_site_code(fields.get("location", DEFAULT_SITE), True)
            if not check_site_permission(user, location):
                self.send_json(403, {"error": f"Tài khoản không có quyền cập nhật biểu mẫu cho trạm {location}."})
                return

            family = fields.get("family", "")

            if family not in {"rfi", "rp"}:

                raise UserError("Hãy chọn nhóm biểu mẫu RFI hoặc RP.")

            if len(uploads) != 1:

                raise UserError("Mỗi lần chỉ tải lên một biểu mẫu.")

            filename, content = uploads[0]

            report_type = validate_template_upload(filename, content, family)

            folder = site_template_dir(location, family)

            folder.mkdir(parents=True, exist_ok=True)

            destination = folder / Path(filename.replace("\\", "/")).name

            if destination.parent.resolve() != folder.resolve():

                raise UserError("Tên file biểu mẫu không hợp lệ.")

            if destination.exists():

                backup_dir = site_data_dir(location) / "Backup" / "Templates" / family

                backup_dir.mkdir(parents=True, exist_ok=True)

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

                shutil.copy2(destination, backup_dir / f"{timestamp}_{destination.name}")

            temporary = destination.with_name(destination.name + ".upload")

            temporary.write_bytes(content)

            temporary.replace(destination)

            apply_station_code_to_template_file(

                destination, station_code_for_location(location)

            )

            try:

                destination.touch()

            except Exception:

                pass

            now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            u_name = str(user.get("name") or user.get("email") or "Kỹ sư")
            save_source_upload_meta(location, f"template_{family}_{report_type.lower()}", filename, len(content), now_str, u_name)

            set_active_site(location)

            for existing in folder.iterdir():

                if existing == destination or existing.suffix.lower() not in {".xlsx", ".xlsm"}:

                    continue

                try:

                    existing_type = template_report_type(existing.name, family)

                except UserError:

                    continue

                if existing_type != report_type:

                    continue

                backup_dir = site_data_dir(location) / "Backup" / "Templates" / family

                backup_dir.mkdir(parents=True, exist_ok=True)

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

                shutil.copy2(existing, backup_dir / f"{timestamp}_{existing.name}")

                existing.unlink()

            self.send_json(

                200,

                {

                    "message": f"Đã cập nhật biểu mẫu {report_type} cho {location}.",

                    "filename": destination.name,

                    "location": location,

                    "family": family,

                },

            )

        except UserError as exc:

            self.send_json(400, {"error": str(exc)})

        except Exception as exc:

            self.send_json(500, {"error": f"Không cập nhật được biểu mẫu: {exc}"})

    def handle_export(self) -> None:

        try:

            content_length = int(self.headers.get("Content-Length", "0"))

            if content_length < 1 or content_length > 2 * 1024 * 1024:

                raise UserError("Danh sách số RFI không hợp lệ.")

            payload = json.loads(self.rfile.read(content_length).decode("utf-8"))

            selected_ids = payload.get("ids")

            if not isinstance(selected_ids, list) or not all(

                isinstance(item, str) for item in selected_ids

            ):

                raise UserError("Danh sách số RFI không hợp lệ.")

            mode = payload.get("mode", "rfi")

            if not isinstance(mode, str):

                raise UserError("Chế độ xuất không hợp lệ.")

            location = validate_site_code(payload.get("location", DEFAULT_SITE), True)

            user = self.get_current_user()
            if not user:
                self.send_json(401, {"error": "Vui lòng đăng nhập để xuất báo cáo."})
                return
            u_role = str(user.get("role", "guest")).lower()
            if u_role == "guest":
                self.send_json(403, {"error": "Tài khoản quyền 'Khách' chỉ có quyền xem dữ liệu và tiến độ, không có quyền xuất hoặc tải báo cáo."})
                return
            ok, err = check_role_permission(user, ("admin", "qc", "construction"))
            if not ok:
                self.send_json(403, {"error": err})
                return
            if not check_site_permission(user, location):
                self.send_json(403, {"error": f"Tài khoản không có quyền xuất báo cáo cho trạm {location}."})
                return

            path, download_name = create_download(selected_ids, mode, location)

            content_type = (

                "application/pdf"

                if path.suffix.lower() == ".pdf"

                else "application/zip"

                if path.suffix.lower() == ".zip"

                else "application/vnd.ms-excel.sheet.macroEnabled.12"

                if path.suffix.lower() == ".xlsm"

                else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            )

            self.send_response(200)

            self.send_header("Content-Type", content_type)

            self.send_header(

                "Content-Disposition",

                f'attachment; filename="{download_name}"; filename*=UTF-8\'\'{quote(download_name)}',

            )

            self.send_header("Content-Length", str(path.stat().st_size))

            self.send_header("Cache-Control", "no-store")

            self.end_headers()

            with path.open("rb") as stream:

                shutil.copyfileobj(stream, self.wfile)

        except (ConnectionResetError, ConnectionAbortedError, BrokenPipeError):

            pass

        except UserError as exc:

            self.send_json(400, {"error": str(exc)})

        except Exception as exc:

            self.send_json(500, {"error": f"Không tạo được biểu mẫu: {exc}"})

    def handle_fix_station_code(self) -> None:
        user = self.get_current_user()
        ok, err = check_role_permission(user, ("admin", "qc"))
        if not ok:
            self.send_json(403, {"error": err})
            return
        try:

            content_length = int(self.headers.get("Content-Length", "0"))

            payload = {}

            if content_length > 0:

                payload = json.loads(self.rfile.read(content_length).decode("utf-8"))

            location = validate_site_code(payload.get("location", DEFAULT_SITE), True)

            custom_code = str(payload.get("station_code", "")).strip()

            res = fix_all_templates_for_site(location, custom_code)

            self.send_json(

                200,

                {

                    "message": f"Đã quét và tự động sửa ký tự vị trí ({res['station_code']}) cho {res['updated_files']}/{res['total_files']} biểu mẫu của vị trí {location}.",

                    "result": res,

                },

            )

        except UserError as exc:

            self.send_json(400, {"error": str(exc)})

        except Exception as exc:

            self.send_json(500, {"error": f"Không thể sửa ký tự biểu mẫu: {exc}"})

    def log_message(self, format: str, *args: object) -> None:

        try:

            msg = format % args

            print(f"[Fome] {self.address_string()} - {msg}")

        except Exception:

            pass

def main() -> None:

    ensure_data_files()

    server = ThreadingHTTPServer((HOST, PORT), FomeHandler)

    display_host = "127.0.0.1" if HOST == "0.0.0.0" else HOST

    address = f"http://{display_host}:{PORT}/"

    print("Fome local web is running. Keep this window open while using it.")

    print(f"Open: {address}")

    threading.Timer(1.0, lambda: webbrowser.open(address)).start()

    try:

        server.serve_forever()

    except KeyboardInterrupt:

        print("\nStopping Fome web.")

    finally:

        server.server_close()

if __name__ == "__main__":

    main()

