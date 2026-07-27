#!/usr/bin/env python3
"""
PDF -> Excel extractor for MOVE logistics-center brochures.

Each PDF page is a single (scanned / image-only) property sheet that follows a
fixed template:

    +-------------------------------------------------------------+
    | N. <창고명> (상온)                                           |
    | [building photo]   Location(map)        Site Plan           |
    | General Information            |  Space Availability table   |
    |   소재지 / 인근IC / 건폐율 ... |  (공실현황)                  |
    +-------------------------------------------------------------+

Because the pages carry no text layer, every value is recovered with OCR
(Tesseract, Korean + English). The General-Information table is located by
detecting its horizontal grid lines, then each value cell is OCR'd on its own
which is far more reliable than reading the whole block at once.

The script writes one Excel row per page with these columns:

    사용가능여부 · 중개인/임대인명 · 창고명 · 주소 · 행정구역_도 · 행정구역_시
    · 대지면적 · 연면적 · 건축면적 · 준공연도 · 공실현황(image) · 시설특이사항
    · 기타 · 정보확인일자 · 사진(image)

Usage:
    python main.py dsvtest2.pdf                       # -> dsvtest2.xlsx
    python main.py input.pdf -o out.xlsx --broker S1
    python main.py input.pdf --dpi 300
"""
from __future__ import annotations

import argparse
import io
import os
import re
import shutil
import sys
from copy import deepcopy
from datetime import date
from io import BytesIO

import fitz  # PyMuPDF
import numpy as np
import pytesseract
try:
    import xlsxwriter
except ImportError as exc:
    raise ImportError(
        "XlsxWriter is required to write the Excel output. "
        "Activate the project virtual environment and install it with: "
        "python -m pip install XlsxWriter"
    ) from exc
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image
from templates import TEMPLATES, DEFAULT_TEMPLATE

# ---------------------------------------------------------------------------
# Layout constants (fractions of page width / height -> resolution independent)
# ---------------------------------------------------------------------------
OCR_LANG = "kor+eng"

# General-Information table search box (left half of the lower page).
GI_BOX = dict(x0=0.060, x1=0.302, y0=0.596, y1=0.880)
# x where the value column starts inside a GI row (label sits to the left).
GI_VALUE_X = 0.133
# Ordered field labels of the 9 GI rows.
GI_FIELDS = ["소재지", "인근IC", "건폐율/용적률", "대지면적",
             "연면적", "규모", "형태", "주차", "준공년도"]

# Common label synonyms that appear in some broker templates -> canonical GI_FIELDS
GI_SYNONYMS = {
    "주소": "소재지",
    "위치": "소재지",
    "소재지위치": "소재지",
    "주소및위치": "소재지",
    "준공시기": "준공년도",
    "준공연도": "준공년도",
    "준공일": "준공년도",
    "주차대수": "주차",
    "총연면적": "연면적",
    "연면적㎡": "연면적",
    "연면적m2": "연면적",
    "연면적평": "연면적",
}

# Title band (the "N. 창고명 (상온)" headline).
TITLE_BOX = dict(x0=0.056, x1=0.485, y0=0.142, y1=0.190)

# Building photo region (top-left), trimmed to its real content afterwards.
PHOTO_BOX = dict(x0=0.057, x1=0.306, y0=0.224, y1=0.572)

# Space-Availability table (공실현황) on the right.
SPACE_BOX = dict(x0=0.497, x1=0.978, y0=0.577, y1=0.893)

# Facility-spec / 비고 middle column (시설특이사항).
SPEC_BOX = dict(x0=0.304, x1=0.496, y0=0.603, y1=0.873)

# Canonical province name -> surface forms that may appear at the start of an
# address (longest first so the most specific prefix is stripped).
PROVINCE = [
    ("서울특별시", ["서울특별시", "서울"]),
    ("부산광역시", ["부산광역시", "부산"]),
    ("대구광역시", ["대구광역시", "대구"]),
    ("인천광역시", ["인천광역시", "인천"]),
    ("광주광역시", ["광주광역시", "광주"]),
    ("대전광역시", ["대전광역시", "대전"]),
    ("울산광역시", ["울산광역시", "울산"]),
    ("세종특별자치시", ["세종특별자치시", "세종시", "세종"]),
    ("경기도", ["경기도", "경기"]),
    ("강원특별자치도", ["강원특별자치도", "강원도", "강원"]),
    ("충청북도", ["충청북도", "충북"]),
    ("충청남도", ["충청남도", "충남"]),
    ("전북특별자치도", ["전북특별자치도", "전라북도", "전북"]),
    ("전라남도", ["전라남도", "전남"]),
    ("경상북도", ["경상북도", "경북"]),
    ("경상남도", ["경상남도", "경남"]),
    ("제주특별자치도", ["제주특별자치도", "제주도", "제주"]),
]

COLUMNS = ["사용가능여부", "중개인/임대인명", "창고명", "주소", "행정구역_도",
           "행정구역_시", "대지면적", "연면적", "건축면적", "준공연도",
           "공실현황", "시설특이사항", "기타", "정보확인일자", "사진"]
IMG_COLS = {"공실현황", "사진"}


# ---------------------------------------------------------------------------
# Geometry helpers
# ---------------------------------------------------------------------------
def _px(box: dict, w: int, h: int) -> tuple[int, int, int, int]:
    """Convert a fractional box into absolute pixel coordinates."""
    return (int(box["x0"] * w), int(box["y0"] * h),
            int(box["x1"] * w), int(box["y1"] * h))


def _rect_from_frac(box: dict, page: "fitz.Page") -> "fitz.Rect":
    return fitz.Rect(box["x0"] * page.rect.width, box["y0"] * page.rect.height,
                     box["x1"] * page.rect.width, box["y1"] * page.rect.height)


def render_page(page: "fitz.Page", dpi: int) -> Image.Image:
    pix = page.get_pixmap(dpi=dpi)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


# ---------------------------------------------------------------------------
# OCR helpers
# ---------------------------------------------------------------------------
def ocr(img: Image.Image, psm: int = 6, lang: str = OCR_LANG, scale: int = 1) -> str:
    if scale != 1:
        img = img.resize((img.width * scale, img.height * scale))
    return pytesseract.image_to_string(img, lang=lang, config=f"--psm {psm}").strip()


def get_template(template_name: str) -> dict:
    return TEMPLATES.get(template_name, TEMPLATES[DEFAULT_TEMPLATE])


def detect_grid_rows(gray: np.ndarray, box_px: tuple[int, int, int, int],
                     dark_cut: int = 180, frac: float = 0.5) -> list[int]:
    """Return y-centres of horizontal grid lines inside the GI table box."""
    x0, y0, x1, y1 = box_px
    sub = gray[y0:y1, x0:x1]
    darkness = (sub < dark_cut).mean(axis=1)
    hits = [y0 + y for y, v in enumerate(darkness) if v > frac]
    groups: list[list[int]] = []
    for y in hits:
        if groups and y - groups[-1][-1] <= 4:
            groups[-1].append(y)
        else:
            groups.append([y])
    return [int(np.mean(g)) for g in groups]


# ---------------------------------------------------------------------------
# Value cleaners
# ---------------------------------------------------------------------------
def _strip_noise(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    # OCR frequently appends the right cell border as stray glyphs.
    return text.strip(" :：|.,'\"`~^_-7ㆍ·")


_NUM = r"\d{1,3}(?:,\d{3})*(?:\.\d+)?"


def clean_area(text: str) -> str:
    """Normalise an area string like '29,70501 (8,986평)' -> '29,705㎡ (8,986평)'.

    The ㎡ figure and the (평) figure are read separately with a strict
    thousands/decimal pattern so OCR'd unit glyphs ('m', 'ni', '01', ...) that
    get fused onto the number are discarded.
    """
    left, _, right = text.partition("(")
    sqm_m = re.search(_NUM, left) or re.search(_NUM, text)
    pyeong_m = re.search(rf"({_NUM})\s*평", text) or re.search(_NUM, right)
    if sqm_m and pyeong_m:
        sqm = sqm_m.group(0).rstrip(".,")
        pyeong = (pyeong_m.group(1) if pyeong_m.re.groups else pyeong_m.group(0)).rstrip(".,")
        return f"{sqm}㎡ ({pyeong}평)"
    if sqm_m:
        return f"{sqm_m.group(0).rstrip('.,')}㎡"
    return _strip_noise(text)


def clean_year(text: str) -> str:
    m = re.search(r"((?:19|20)\d{2})\s*년\s*(\d{1,2})?\s*월?", text)
    if m:
        return f"{m.group(1)}년 {m.group(2)}월" if m.group(2) else f"{m.group(1)}년"
    m = re.search(r"(?:19|20)\d{2}", text)
    return f"{m.group(0)}년" if m else ""


def clean_title(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"^\s*\d+\s*[.)]\s*", "", text)          # leading "1. "
    text = re.sub(r"\b[Nn][Ee][Ww]\b", "", text)            # red "NEW" badge
    text = text.strip(" .,")
    # Repair an unterminated "(상온" parenthesis from OCR.
    if text.count("(") > text.count(")"):
        text += ")"
    # Drop a parenthetical that OCR turned into pure garbage (no Hangul inside),
    # e.g. "물류센터 (%)" -> "물류센터".
    text = re.sub(r"\s*\((?![^)]*[가-힣])[^)]*\)", "", text)
    return text.strip()


# Label normalization helpers used by both text- and image-based parsers
def _norm_label(s: str) -> str:
    s = re.sub(r"\s+", "", (s or "")).strip().lower()
    s = s.replace("㎡", "").replace("m2", "")
    s = s.replace("평", "").replace("：", ":").replace(";", "")
    return s


def canonical_field(label: str) -> str | None:
    nl = _norm_label(label)
    # reject overly short noisy OCR tokens (e.g. '재', '면')
    if not nl or len(nl) < 2:
        return None
    # exact synonym matches (trusted)
    for syn, canon in GI_SYNONYMS.items():
        if nl == _norm_label(syn):
            return canon
    # allow matching against canonical field names (including a few S1 extras)
    extra_fields = ["주용도", "구조", "층고"]
    all_fields = GI_FIELDS + extra_fields
    for fld in all_fields:
        fn = _norm_label(fld)
        # exact match or canonical name appearing inside a noisy OCR label
        if nl == fn or fn in nl:
            return fld
    return None


def clean_special_note(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"^비\s*고\s*", "", text)
    return text.strip()


def _split_korean_address(address: str) -> tuple[str, str, str]:
    """Split a Korean address into (도, 시/군, rest) from compact or spaced text."""
    addr = re.sub(r"\s+", "", address.strip())
    do = ""
    for canonical, forms in PROVINCE:
        matched = next((f for f in forms if addr.startswith(f)), None)
        if matched:
            do, addr = canonical, addr[len(matched):]
            break
    si = ""
    m = re.match(r"([가-힣]+?(?:시|군|구))", addr)
    if m:
        si = m.group(1)
        addr = addr[m.end():]
    return do, si, addr


def clean_address(address: str) -> str:
    text = _strip_noise(address)
    if not text:
        return ""
    do, si, rest = _split_korean_address(text)
    rest = re.sub(r"(?<=[가-힣])(?=\d)", " ", rest)
    rest = re.sub(r"(?<=[시군구읍면동])(?=[가-힣])", " ", rest)
    parts = [part for part in (do, si) if part]
    if rest:
        parts.append(rest.strip())
    return " ".join(parts).strip()


def parse_region(address: str) -> tuple[str, str]:
    """Split an address into (도, 시/군). Handles spaced and unspaced forms."""
    addr = re.sub(r"\s+", "", address.strip())
    do = ""
    for canonical, forms in PROVINCE:
        matched = next((f for f in forms if addr.startswith(f)), None)
        if matched:
            do, addr = canonical, addr[len(matched):]
            break
    si = ""
    m = re.search(r"([가-힣]+?(?:시|군|구))", addr)
    if m:
        si = m.group(1)
    return do, si


def _extract_gi_values_from_words(page: "fitz.Page", template: dict) -> dict[str, str]:
    """Coordinate-based GI parser for text-layer PDFs.

    Returns a dict mapping canonical GI fields to their extracted values.
    """
    values: dict[str, str] = {}
    box = template.get("GI_BOX", GI_BOX)
    rect = fitz.Rect(box["x0"] * page.rect.width, box["y0"] * page.rect.height,
                     box["x1"] * page.rect.width, box["y1"] * page.rect.height)
    value_x = template.get("GI_VALUE_X", GI_VALUE_X) * page.rect.width

    words = page.get_text("words")
    groups: dict[int, list[tuple[float, str]]] = {}
    for w in words:
        cx = (w[0] + w[2]) / 2.0
        cy = (w[1] + w[3]) / 2.0
        if not rect.contains(fitz.Point(cx, cy)):
            continue
        yk = int(cy)
        groups.setdefault(yk, []).append((cx, w[4]))

    # merge nearby Y rows
    rows: list[tuple[int, list[tuple[float, str]]]] = []
    for y in sorted(groups.keys()):
        if rows and y - rows[-1][0] <= 6:
            rows[-1][1].extend(groups[y])
        else:
            rows.append((y, list(groups[y])))

    gi_fields = template.get("GI_FIELDS", GI_FIELDS)
    row_values: list[tuple[str, str]] = []
    for _, items in rows:
        items_sorted = [t for _, t in sorted(items, key=lambda x: x[0])]
        # split left/right by x using the original word positions
        left_parts: list[str] = []
        right_parts: list[str] = []
        for cx, txt in sorted(items, key=lambda x: x[0]):
            if cx < value_x:
                left_parts.append(txt)
            else:
                right_parts.append(txt)
        left_label = "".join(left_parts).replace(" ", "")
        right_val = " ".join(right_parts).strip()
        row_values.append((left_label, right_val))

    # First pass: resolve labels explicitly
    for left_label, right_val in row_values:
        fld = canonical_field(left_label)
        if fld:
            values.setdefault(fld, right_val)

    # Second pass: fallback to positional mapping using template GI_FIELDS
    if len(values) < max(3, len(gi_fields) // 2):
        for idx, (left_label, right_val) in enumerate(row_values):
            if idx >= len(gi_fields):
                break
            fld = gi_fields[idx]
            values.setdefault(fld, right_val)

    return values


# ---------------------------------------------------------------------------
# Image cropping
# ---------------------------------------------------------------------------
def content_bbox(img: Image.Image, box_px, white: int = 235, pad: int = 6):
    """Tighten a box to the non-white content it contains."""
    x0, y0, x1, y1 = box_px
    sub = np.array(img.crop((x0, y0, x1, y1)).convert("L"))
    mask = sub < white
    ys = np.where(mask.any(axis=1))[0]
    xs = np.where(mask.any(axis=0))[0]
    if len(xs) == 0 or len(ys) == 0:
        return box_px
    return (max(x0 + int(xs.min()) - pad, 0), max(y0 + int(ys.min()) - pad, 0),
            x0 + int(xs.max()) + pad, y0 + int(ys.max()) + pad)


# ---------------------------------------------------------------------------
# Per-page extraction
# ---------------------------------------------------------------------------
def extract_page(img: Image.Image, page_idx: int, img_dir: str,
                 template: dict, crop_dpi: int = 150) -> dict:
    w, h = img.size
    gray = np.array(img.convert("L"))
    gi_box = _px(template.get("GI_BOX", GI_BOX), w, h)
    rows = detect_grid_rows(gray, gi_box)
    values: dict[str, str] = {}
    gi_fields = template.get("GI_FIELDS", GI_FIELDS)
    val_x0 = int(template.get("GI_VALUE_X", GI_VALUE_X) * w)
    val_x1 = gi_box[2]
    label_x0 = gi_box[0]
    label_x1 = val_x0
    # If grid detection failed or found too few rows, fallback to block OCR
    if len(rows) < 4:
        gi_img = img.crop(gi_box)
        gi_text = ocr(gi_img, psm=6, lang=template.get("OCR_LANG", OCR_LANG), scale=2)
        lines = [ln.strip() for ln in gi_text.splitlines() if ln.strip()]
        j = 0
        while j < len(lines):
            ln = lines[j]
            m = re.match(r"^([^:：\t]+?)\s*[:：]\s*(.+)$", ln)
            if m:
                lab, val = m.group(1).strip(), m.group(2).strip()
                fld = canonical_field(lab)
                if fld:
                    values.setdefault(fld, _strip_noise(val))
            else:
                fld = canonical_field(ln)
                if fld and j + 1 < len(lines):
                    values.setdefault(fld, _strip_noise(lines[j + 1].strip()))
                    j += 1
                else:
                    parts = re.split(r"\s{2,}", ln)
                    if len(parts) >= 2:
                        fld = canonical_field(parts[0].strip())
                        if fld:
                            values.setdefault(fld, _strip_noise(parts[1].strip()))
            j += 1
    else:
        for i in range(len(rows) - 1):
            if i >= len(gi_fields):
                break
            yt, yb = rows[i] + 2, rows[i + 1] - 2
            if yb - yt < 8:
                continue
            # OCR the label area to infer which canonical field this row is
            try:
                label_img = img.crop((label_x0, yt, label_x1, yb))
                label_txt = ocr(label_img, psm=7, lang=template.get("OCR_LANG", OCR_LANG), scale=2)
                fld = canonical_field(label_txt)
            except Exception:
                fld = None

            cell = img.crop((val_x0, yt, val_x1, yb))
            val_txt = _strip_noise(ocr(cell, psm=7, lang=template.get("OCR_LANG", OCR_LANG), scale=2))
            target_field = fld if fld else gi_fields[i]
            # If the address value is suspiciously short, OCR a wider label+value area
            if target_field == "소재지" and (not val_txt or len(val_txt.strip()) < 4):
                try:
                    wide = img.crop((label_x0, max(0, yt - 4), val_x1, min(h, yb + 6)))
                    wide_txt = _strip_noise(ocr(wide, psm=6, lang=template.get("OCR_LANG", OCR_LANG), scale=2))
                    if wide_txt:
                        values[target_field] = wide_txt
                    else:
                        values[target_field] = val_txt
                except Exception:
                    values[target_field] = val_txt
            else:
                values[target_field] = val_txt

    address = clean_address(values.get("소재지", ""))
    do, si = parse_region(address)
    year = clean_year(values.get("준공년도", "")) or "미정"

    # Heuristic fallback: if critical fields look wrong (very short or placeholder
    # OCR tokens), re-ocr the whole GI box and try to extract missing fields.
    def _looks_bad_address(s: str) -> bool:
        if not s:
            return True
        s2 = s.strip()
        return len(s2) < 4 or s2 in {"재", "면", "-"}

    def _looks_bad_area(s: str) -> bool:
        if not s:
            return True
        return not re.search(r"\d", s)

    if _looks_bad_address(values.get("소재지", "")) or _looks_bad_area(values.get("연면적", "")):
        try:
            gi_img = img.crop(gi_box)
            gi_text = ocr(gi_img, psm=6, lang=template.get("OCR_LANG", OCR_LANG), scale=2)
            lines = [ln.strip() for ln in gi_text.splitlines() if ln.strip()]
            for ln in lines:
                m = re.match(r"^([^:：\t]+?)\s*[:：]\s*(.+)$", ln)
                if m:
                    lab, val = m.group(1).strip(), m.group(2).strip()
                    fld = canonical_field(lab)
                    if fld:
                        # only replace if missing or suspicious
                        if fld == "소재지":
                            if _looks_bad_address(values.get(fld, "")):
                                values[fld] = _strip_noise(val)
                        elif fld == "연면적":
                            if _looks_bad_area(values.get(fld, "")):
                                values[fld] = _strip_noise(val)
                        else:
                            values.setdefault(fld, _strip_noise(val))
                else:
                    # Catch lines like '대지면적 53,945㎡ (16,318평)'
                    parts = re.split(r"\s{2,}", ln)
                    if len(parts) >= 2:
                        lab_cand, val_cand = parts[0].strip(), parts[1].strip()
                        fld = canonical_field(lab_cand)
                        if fld:
                            if fld == "소재지" and _looks_bad_address(values.get(fld, "")):
                                values[fld] = _strip_noise(val_cand)
                            elif fld == "연면적" and _looks_bad_area(values.get(fld, "")):
                                values[fld] = _strip_noise(val_cand)
            # numeric area fallback search
            if _looks_bad_area(values.get("연면적", "")):
                mnum = re.search(_NUM + r"\s*(?:㎡|m2|㎡)?", gi_text)
                if mnum:
                    values["연면적"] = f"{mnum.group(0).rstrip('.,')}㎡"
            # address fallback: find a province + city fragment inside block text
            if _looks_bad_address(values.get("소재지", "")):
                for canonical, forms in PROVINCE:
                    for f in forms:
                        idx = gi_text.find(f)
                        if idx != -1:
                            # take substring from match to end of line
                            tail = gi_text[idx:]
                            tail_line = tail.splitlines()[0]
                            values["소재지"] = _strip_noise(tail_line)
                            break
                    if values.get("소재지"):
                        break
        except Exception:
            pass

    # recompute cleaned address/region/year after fallback
    address = clean_address(values.get("소재지", ""))
    do, si = parse_region(address)
    year = clean_year(values.get("준공년도", "")) or "미정"

    # --- Title (창고명) ----------------------------------------------------
    title_box = _px(template.get("TITLE_BOX", TITLE_BOX), w, h)
    title = clean_title(ocr(img.crop(title_box), psm=7,
                            lang=template.get("OCR_LANG", OCR_LANG), scale=2))

    # --- 시설특이사항 (facility spec / 비고 cell) ------------------------
    spec = ""
    spec_box = template.get("SPEC_BOX", SPEC_BOX)
    try:
        spec_img = img.crop(_px(spec_box, w, h))
        spec = clean_special_note(ocr(spec_img, psm=6,
                                       lang=template.get("OCR_LANG", OCR_LANG),
                                       scale=2))
    except Exception:
        spec = ""

    # --- Images: building photo + Space-Availability table ----------------
    photo_path = os.path.join(img_dir, f"photo_p{page_idx + 1}.png")
    try:
        photo_box_px = _px(template.get("PHOTO_BOX", PHOTO_BOX), w, h)
        img.crop(photo_box_px).save(photo_path)
    except Exception:
        pass

    space_path = os.path.join(img_dir, f"space_p{page_idx + 1}.png")
    try:
        space_box_px = _px(template.get("SPACE_BOX", SPACE_BOX), w, h)
        # apply template-specific SPACE_PAD when present, otherwise fall back
        # to the historic small expansion that helped include headers.
        lx0, ly0, lx1, ly1 = space_box_px
        space_pad = template.get("SPACE_PAD")
        if space_pad is not None:
            lx0 = max(0, lx0 - int(space_pad.get("left", 0.0) * w))
            ly0 = max(0, ly0 - int(space_pad.get("top", 0.0) * h))
            lx1 = min(w, lx1 + int(space_pad.get("right", 0.0) * w))
            ly1 = min(h, ly1 + int(space_pad.get("bottom", 0.0) * h))
        else:
            lx0 = max(0, lx0 - int(0.05 * w))
            ly0 = max(0, ly0 - int(0.03 * h))
            lx1 = min(w, lx1 + int(0.01 * w))
            ly1 = min(h, ly1 + int(0.02 * h))
        img.crop((lx0, ly0, lx1, ly1)).save(space_path)
    except Exception:
        pass

    return {
        "창고명": title,
        "주소": address,
        "행정구역_도": do,
        "행정구역_시": si,
        "대지면적": clean_area(values.get("대지면적", "")),
        "연면적": clean_area(values.get("연면적", "")),
        "준공연도": year,
        "시설특이사항": spec,
        "_photo": photo_path,
        "_space": space_path,
    }


# ---------------------------------------------------------------------------
# Text-layer extraction (for PDFs that already carry selectable text)
# ---------------------------------------------------------------------------
def is_property_page(text: str, template: dict) -> bool:
    """Detect a property sheet page using the selected template's keywords."""
    text_norm = re.sub(r"\s+", "", text)
    keywords = template.get("keywords")
    if keywords:
        found = sum(1 for keyword in keywords
                    if re.sub(r"\s+", "", keyword) in text_norm)
        if found < max(1, len(keywords) // 2):
            return False
    if re.search(r"\d+\.\s*\S", text):
        return True
    return False


def _photo_rect(page: "fitz.Page"):
    """The building render is the largest embedded image in the top-left."""
    best, best_area = None, 0.0
    for im in page.get_images(full=True):
        for r in page.get_image_rects(im[0]):
            if r.x1 < 255 and r.y0 < 205:
                area = r.width * r.height
                if area > best_area and area > 8000:
                    best, best_area = r, area
    if best is None:  # fallback to a fixed left-column box
        best = fitz.Rect(32, 122, 233, 312)
    return best


def _space_rect(page: "fitz.Page"):
    """Box around the Space-Availability table: header down to the 합계 row."""
    sx = sy = total_bottom = None
    for w in page.get_text("words"):
        if w[4] == "Space":
            sx, sy = w[0], w[1]
        if w[4] == "합계":
            total_bottom = w[3]
    left = (sx - 6) if sx is not None else 388
    top = (sy - 8) if sy is not None else 310
    bottom = (total_bottom + 9) if total_bottom is not None else 495
    return fitz.Rect(left, top, page.rect.width - 6, bottom)


def _render_clip(page: "fitz.Page", rect, dpi: int, out_path: str) -> None:
    m = fitz.Matrix(dpi / 72, dpi / 72)
    page.get_pixmap(matrix=m, clip=rect).save(out_path)


def _words_in_box_lines(page: "fitz.Page", box: dict) -> list[str]:
    """Return ordered text lines from words that fall inside a fractional box."""
    r = fitz.Rect(box["x0"] * page.rect.width, box["y0"] * page.rect.height,
                  box["x1"] * page.rect.width, box["y1"] * page.rect.height)
    words = page.get_text("words")
    groups: dict[int, list[tuple[float, str]]] = {}
    for w in words:
        wr = fitz.Rect(w[0], w[1], w[2], w[3])
        if not wr.intersects(r):
            continue
        y = int(w[1])
        groups.setdefault(y, []).append((w[0], w[4]))
    lines_out: list[str] = []
    for y in sorted(groups.keys()):
        parts = [t for _, t in sorted(groups[y], key=lambda x: x[0])]
        lines_out.append(" ".join(parts).strip())
    return lines_out


def _extract_spec_from_box(page: "fitz.Page", template: dict) -> str:
    spec_box = template.get("SPEC_BOX", SPEC_BOX)
    lines = _words_in_box_lines(page, spec_box)
    normalized: list[str] = []
    i = 0
    while i < len(lines):
        if lines[i] == "비" and i + 1 < len(lines) and lines[i + 1] == "고":
            normalized.append("비고")
            i += 2
            continue
        normalized.append(lines[i])
        i += 1

    start = next((idx for idx, ln in enumerate(normalized) if ln == "비고"), None)
    if start is None:
        return "\n".join(normalized).strip()

    spec_lines = []
    stop_tokens = set(GI_FIELDS) | {
        "본부장", "부장", "대리", "Location", "층", "총임대면적(평)",
        "임대가능면적(평)", "임대가능시점", "비고", "구", "분", "내", "용"
    }
    for ln in normalized[start + 1:]:
        if ln in stop_tokens:
            break
        spec_lines.append(ln)
    return "\n".join(spec_lines).strip()


def _is_valid_title(title: str, template: dict) -> bool:
    title = (title or "").strip()
    if not title:
        return False
    if any(token in title for token in ("%", "㎡", "평", "용적률", "건폐율")):
        return False
    compact = re.sub(r"\s+", "", title)
    if len(compact) < 4:
        return False
    if not re.search(r"[A-Za-z가-힣]", title):
        return False
    required = template.get("TITLE_REQUIRED_TOKENS", [])
    if required and not any(token in title for token in required):
        return False
    return True


def _words_centered_in_box_lines(page: "fitz.Page", box: dict) -> list[str]:
    rect = _rect_from_frac(box, page)
    words = page.get_text("words")
    rows: dict[int, list[tuple[float, str]]] = {}
    for w in words:
        cx = (w[0] + w[2]) / 2.0
        cy = (w[1] + w[3]) / 2.0
        if not rect.contains(fitz.Point(cx, cy)):
            continue
        key = int(cy // 6)
        rows.setdefault(key, []).append((cx, w[4]))
    lines: list[str] = []
    for _, items in sorted(rows.items()):
        items_sorted = [text for _, text in sorted(items, key=lambda x: x[0])]
        lines.append(" ".join(items_sorted).strip())
    return lines


def _title_from_top_region(page: "fitz.Page", template: dict) -> str:
    top_threshold = page.rect.height * 0.15
    words = page.get_text("words")
    rows: dict[int, list[tuple[float, str]]] = {}
    for w in words:
        cx = (w[0] + w[2]) / 2.0
        cy = (w[1] + w[3]) / 2.0
        if cy > top_threshold:
            continue
        key = int(cy // 6)
        rows.setdefault(key, []).append((cx, w[4]))
    for _, items in sorted(rows.items()):
        line = " ".join(text for _, text in sorted(items, key=lambda x: x[0])).strip()
        if _is_valid_title(line, template):
            return line
    return ""


def _select_title_from_box(page: "fitz.Page", template: dict, page_text: str) -> str:
    """Pick a likely title from the TITLE_BOX words, fall back to page_text scan."""
    try:
        box = template.get("TITLE_BOX", TITLE_BOX)
        candidates = _words_centered_in_box_lines(page, box)
        required = template.get("TITLE_REQUIRED_TOKENS", [])
        for ln in candidates:
            if not ln:
                continue
            if not _is_valid_title(ln, template):
                continue
            return clean_title(ln)
        if required:
            return ""
        for ln in candidates:
            if not ln:
                continue
            if _is_valid_title(ln, template):
                return clean_title(ln)
    except Exception:
        pass
    return _title_from_top_region(page, template)


def extract_text_page(page: "fitz.Page", page_idx: int,
                      img_dir: str, crop_dpi: int, template: dict) -> dict:
    """Parse one property sheet from its text layer + render the two images.

    Uses the template `TITLE_BOX` region to pick a reliable title when
    available, otherwise falls back to heuristics over the page text.
    """
    text = page.get_text()
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    # Attempt a focused extraction of the title from the template TITLE_BOX
    title = _select_title_from_box(page, template, text)

    # If the selectable title is invalid, avoid arbitrary fallback lines.
    if not _is_valid_title(title, template):
        title = ""

    # If box-based selection failed, try the top region of the page with selectable
    # text only, and never fall back to broad heuristic line scanning.
    if not title:
        title = _title_from_top_region(page, template)

    # Coordinate-based GI extraction first (uses visual positions inside the
    # template GI_BOX). Fall back to the loose line parser below only when
    # the coordinate-based extraction doesn't provide values.
    values: dict[str, str] = _extract_gi_values_from_words(page, template)

    def _norm(s: str) -> str:
        return _norm_label(s)

    def match_field(label: str) -> str | None:
        return canonical_field(label)

    i = 0
    while i < len(lines):
        ln = lines[i]
        m = re.match(r"^\s*([^:：\t]+?)\s*[:：]\s*(.+)$", ln)
        if m:
            lab, val = m.group(1).strip(), m.group(2).strip()
            fld = match_field(lab)
            if fld:
                values.setdefault(fld, val)
        else:
            fld = match_field(ln)
            if fld and i + 1 < len(lines):
                values.setdefault(fld, lines[i + 1].strip())
                i += 1
            else:
                parts = re.split(r"\s{2,}", ln)
                if len(parts) >= 2:
                    lab_cand, val_cand = parts[0].strip(), parts[1].strip()
                    fld = match_field(lab_cand)
                    if fld:
                        values.setdefault(fld, val_cand)
        i += 1

    spec = ""
    if template.get("SPEC_BOX"):
        spec = _extract_spec_from_box(page, template)

    if not spec:
        # Normalize consecutive '비' + '고' lines into a single '비고' line.
        normalized_lines: list[str] = []
        j = 0
        while j < len(lines):
            if lines[j] == "비" and j + 1 < len(lines) and lines[j + 1] == "고":
                normalized_lines.append("비고")
                j += 2
            else:
                normalized_lines.append(lines[j])
                j += 1

        bi_index = next((idx for idx, ln in enumerate(normalized_lines)
                         if _norm(ln).startswith(_norm("비고"))), None)
        if bi_index is not None:
            spec_lines = []
            for ln in normalized_lines[bi_index + 1:]:
                if match_field(ln) or _norm(ln) in (_norm("본부장"), _norm("부장"), _norm("대리"), _norm("Location")):
                    break
                spec_lines.append(ln)
            spec = "\n".join(spec_lines)

    if not spec:
        try:
            spec_lines = _words_in_box_lines(page, template.get("SPEC_BOX", SPEC_BOX))
            spec_lines = [ln for ln in spec_lines if not any(_norm(ln).startswith(_norm(x)) for x in GI_FIELDS)]
            spec = "\n".join(spec_lines).strip()
        except Exception:
            pass

    address = clean_address(values.get("소재지", ""))
    do, si = parse_region(address)
    year = clean_year(values.get("준공년도", "")) or "미정"

    photo_path = os.path.join(img_dir, f"photo_p{page_idx + 1}.png")
    # Prefer the template PHOTO_BOX when available (text PDF pages)
    try:
        if template.get("PHOTO_BOX"):
            _render_clip(page, _rect_from_frac(template.get("PHOTO_BOX"), page), crop_dpi, photo_path)
        else:
            _render_clip(page, _photo_rect(page), crop_dpi, photo_path)
    except Exception:
        _render_clip(page, _photo_rect(page), crop_dpi, photo_path)

    space_path = os.path.join(img_dir, f"space_p{page_idx + 1}.png")
    # For templates with a known SPACE_BOX, render that region; apply template
    # SPACE_PAD when provided so S1 can avoid left expansion into 기타사항.
    try:
        box = template.get("SPACE_BOX")
        pad = template.get("SPACE_PAD")
        if box:
            if pad is not None:
                adj = {
                    "x0": max(0.0, box.get("x0", 0.0) - pad.get("left", 0.0)),
                    "y0": max(0.0, box.get("y0", 0.0) - pad.get("top", 0.0)),
                    "x1": min(1.0, box.get("x1", 1.0) + pad.get("right", 0.0)),
                    "y1": min(1.0, box.get("y1", 1.0) + pad.get("bottom", 0.0)),
                }
                _render_clip(page, _rect_from_frac(adj, page), crop_dpi, space_path)
            else:
                _render_clip(page, _rect_from_frac(box, page), crop_dpi, space_path)
        else:
            _render_clip(page, _space_rect(page), crop_dpi, space_path)
    except Exception:
        _render_clip(page, _space_rect(page), crop_dpi, space_path)

    return {
        "창고명": title,
        "주소": address,
        "행정구역_도": do,
        "행정구역_시": si,
        "대지면적": clean_area(values.get("대지면적", "")),
        "연면적": clean_area(values.get("연면적", "")),
        "준공연도": year,
        "시설특이사항": spec,
        "_photo": photo_path,
        "_space": space_path,
    }


# ---------------------------------------------------------------------------
# Excel writing
# ---------------------------------------------------------------------------
def write_excel(records: list[dict], out_path: str, broker: str,
                info_date: str, img_target_w: int) -> None:
    workbook = xlsxwriter.Workbook(out_path)
    ws = workbook.add_worksheet("Warehouses")

    header_format = workbook.add_format({
        "bold": True,
        "align": "center",
        "valign": "vcenter",
        "text_wrap": True,
        "fg_color": "#2F5597",
        "font_color": "#FFFFFF",
        "border": 1,
    })
    cell_format = workbook.add_format({
        "align": "left",
        "valign": "top",
        "text_wrap": True,
        "border": 1,
    })

    for c, name in enumerate(COLUMNS):
        ws.write(0, c, name, header_format)
    ws.set_row_pixels(0, 24)

    col_index = {name: i + 1 for i, name in enumerate(COLUMNS)}
    default_widths = {
        "사용가능여부": 12, "중개인/임대인명": 14, "창고명": 26, "주소": 26,
        "행정구역_도": 14, "행정구역_시": 12, "대지면적": 20, "연면적": 20,
        "건축면적": 14, "준공연도": 12, "시설특이사항": 38, "기타": 12,
        "정보확인일자": 14,
    }
    for name, width in default_widths.items():
        ws.set_column(col_index[name] - 1, col_index[name] - 1, width)

    image_columns = ["공실현황", "사진"]
    for name in image_columns:
        ws.set_column_pixels(col_index[name] - 1, col_index[name] - 1, img_target_w)

    for row_idx, rec in enumerate(records, start=1):
        row_values = {
            "사용가능여부": "",
            "중개인/임대인명": broker,
            "창고명": rec["창고명"],
            "주소": rec["주소"],
            "행정구역_도": rec["행정구역_도"],
            "행정구역_시": rec["행정구역_시"],
            "대지면적": rec["대지면적"],
            "연면적": rec["연면적"],
            "건축면적": "",
            "준공연도": rec["준공연도"],
            "공실현황": "",
            "시설특이사항": rec["시설특이사항"],
            "기타": "",
            "정보확인일자": info_date,
            "사진": "",
        }
        max_row_px = 24

        for name, value in row_values.items():
            ws.write(row_idx, col_index[name] - 1, value, cell_format)

        for name, path in (("공실현황", rec["_space"]), ("사진", rec["_photo"])):
            if not path:
                continue

            with Image.open(path) as im:
                ow, oh = im.size
            if ow == 0:
                continue

            scale = img_target_w / ow
            scaled_height = int(oh * scale)
            max_row_px = max(max_row_px, scaled_height + 10)

            ws.embed_image(row_idx, col_index[name] - 1, path,
                           {"description": name, "x_scale": scale, "y_scale": scale})

        ws.set_row_pixels(row_idx, max_row_px)

    ws.freeze_panes(1, 0)
    workbook.close()


def merge_excel_bytes(workbook_bytes: list[bytes]) -> bytes:
    """Merge multiple broker .xlsx workbooks into one combined workbook."""
    if not workbook_bytes:
        raise ValueError("No workbooks provided for merging.")

    wb = Workbook()
    ws = wb.active
    ws.title = "Warehouses"

    header_fill = PatternFill("solid", fgColor="2F5597")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap_top = Alignment(wrap_text=True, vertical="top")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for c, name in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=c, value=name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    default_widths = {
        "사용가능여부": 12, "중개인/임대인명": 14, "창고명": 26, "주소": 26,
        "행정구역_도": 14, "행정구역_시": 12, "대지면적": 20, "연면적": 20,
        "건축면적": 14, "준공연도": 12, "시설특이사항": 38, "기타": 12,
        "정보확인일자": 14,
    }
    col_index = {name: i + 1 for i, name in enumerate(COLUMNS)}
    for name, width in default_widths.items():
        ws.column_dimensions[get_column_letter(col_index[name])].width = width

    row_offset = 2
    for workbook_data in workbook_bytes:
        src_wb = load_workbook(filename=BytesIO(workbook_data), data_only=True)
        src_ws = src_wb.active
        header = [cell.value for cell in src_ws[1]]
        if header != COLUMNS:
            raise ValueError("Cannot merge workbooks with incompatible headers.")

        rows_before = row_offset - 2
        for src_row_idx, row_values in enumerate(src_ws.iter_rows(min_row=2, values_only=True), start=2):
            dst_row = rows_before + src_row_idx
            for col, value in enumerate(row_values, start=1):
                cell = ws.cell(row=dst_row, column=col, value=value)
                cell.alignment = wrap_top
                cell.border = border
            src_dim = src_ws.row_dimensions.get(src_row_idx)
            if src_dim and src_dim.height:
                ws.row_dimensions[dst_row].height = src_dim.height
            row_offset += 1

        for col_letter, dim in src_ws.column_dimensions.items():
            if dim and getattr(dim, "width", None) is not None:
                current_width = ws.column_dimensions[col_letter].width or 0
                ws.column_dimensions[col_letter].width = max(current_width, dim.width)

        for img in getattr(src_ws, "_images", []):
            try:
                anchor = img.anchor
                src_row = anchor._from.row + 1
                src_col = anchor._from.col + 1
                dst_row = rows_before + src_row
                if dst_row >= 2:
                    if hasattr(img.ref, "seek"):
                        img.ref.seek(0)
                    image_bytes = img.ref.read() if hasattr(img.ref, "read") else None
                    merged_image = XLImage(BytesIO(image_bytes) if image_bytes is not None else img.path)

                    try:
                        merged_anchor = TwoCellAnchor(editAs="twoCell")
                        merged_anchor._from = AnchorMarker(col=src_col - 1, row=dst_row - 1, colOff=0, rowOff=0)
                        merged_anchor.to = AnchorMarker(col=src_col, row=dst_row, colOff=0, rowOff=0)
                        merged_image.anchor = merged_anchor
                    except Exception:
                        try:
                            merged_image.anchor = f"{get_column_letter(src_col)}{dst_row}"
                        except Exception:
                            pass

                    ws.add_image(merged_image)
            except Exception:
                continue

    ws.freeze_panes = "A2"
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()


# ---------------------------------------------------------------------------
# Shared extraction driver (used by the CLI and the Streamlit app)
# ---------------------------------------------------------------------------
def ocr_available() -> bool:
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def extract_records(doc: "fitz.Document", img_dir: str, dpi: int = 300,
                    img_dpi: int = 150, ocr_ready: bool | None = None,
                    progress=None, template_name: str = DEFAULT_TEMPLATE) -> tuple[list[dict], int]:
    """Walk every page, returning (property records, skipped page count).

    `progress`, if given, is called as progress(page_index, total, record_or_None).
    """
    if ocr_ready is None:
        ocr_ready = ocr_available()
    os.makedirs(img_dir, exist_ok=True)
    template = get_template(template_name)

    records: list[dict] = []
    skipped = 0
    total = len(doc)
    for i, page in enumerate(doc):
        text = page.get_text()
        rec = None
        if is_property_page(text, template):
            rec = extract_text_page(page, i, img_dir, img_dpi, template)
        elif len(text.strip()) < 50 and ocr_ready:
            rec = extract_page(render_page(page, dpi), i, img_dir, template, crop_dpi=img_dpi)
        else:
            skipped += 1
        if rec is not None:
            records.append(rec)
        if progress is not None:
            progress(i, total, rec)
    return records, skipped


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Extract MOVE logistics-center PDFs into an Excel sheet.")
    parser.add_argument("pdf", help="input PDF path")
    parser.add_argument("-o", "--output", help="output .xlsx path "
                        "(default: <pdf name>.xlsx)")
    parser.add_argument("--dpi", type=int, default=300,
                        help="render DPI for OCR on image-only PDFs (default: 300)")
    parser.add_argument("--img-dpi", type=int, default=300,
                        help="render DPI for cropped cell images (default: 300)")
    parser.add_argument("--broker", default="",
                        help="중개인/임대인명 value for every row (default: empty)")
    parser.add_argument("--date", dest="info_date",
                        help="정보확인일자 value (default: today as dd/mm/yyyy)")
    parser.add_argument("--img-width", type=int, default=360,
                        help="embedded image width in px (default: 360)")
    parser.add_argument("--keep-images", action="store_true",
                        help="keep the cropped image folder (by default it is "
                        "deleted once the images are embedded in the workbook)")
    parser.add_argument("--template", default=DEFAULT_TEMPLATE,
                        help="PDF template to use for extraction")
    args = parser.parse_args(argv)

    if not os.path.isfile(args.pdf):
        print(f"error: file not found: {args.pdf}", file=sys.stderr)
        return 1

    out_path = args.output or os.path.splitext(args.pdf)[0] + ".xlsx"
    info_date = args.info_date or date.today().strftime("%d/%m/%Y")
    img_dir = os.path.splitext(out_path)[0] + "_images"
    os.makedirs(img_dir, exist_ok=True)

    doc = fitz.open(args.pdf)

    def _log(i, total, rec):
        if rec is not None:
            print(f"[page {i + 1}/{total}] {rec['창고명'] or '(no title)'} | "
                  f"{rec['행정구역_도']} {rec['행정구역_시']} | 준공 {rec['준공연도']}")

    records, skipped = extract_records(doc, img_dir, dpi=args.dpi,
                                       img_dpi=args.img_dpi, progress=_log,
                                       template_name=args.template)

    if not records:
        shutil.rmtree(img_dir, ignore_errors=True)
        print("error: no property pages found in this PDF.", file=sys.stderr)
        return 3

    write_excel(records, out_path, args.broker, info_date, args.img_width)
    print(f"\nDone -> {out_path}  ({len(records)} rows, {skipped} non-pattern "
          f"pages skipped)")

    # The images are now embedded in the workbook, so the loose crops and any
    # temporary render are no longer needed.
    if args.keep_images:
        print(f"Cropped images -> {img_dir}/")
    else:
        shutil.rmtree(img_dir, ignore_errors=True)
        print("Cropped images cleaned up (embedded in the workbook).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
