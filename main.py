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
    "준공시기": "준공년도",
    "준공연도": "준공년도",
    "주차대수": "주차",
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
                 template: dict) -> dict:
    w, h = img.size
    gray = np.array(img.convert("L"))
    gi_box = _px(template.get("GI_BOX", GI_BOX), w, h)
    rows = detect_grid_rows(gray, gi_box)
    values: dict[str, str] = {}
    val_x0 = int(template.get("GI_VALUE_X", GI_VALUE_X) * w)
    val_x1 = gi_box[2]
    for i in range(len(rows) - 1):
        if i >= len(GI_FIELDS):
            break
        yt, yb = rows[i] + 2, rows[i + 1] - 2
        if yb - yt < 8:
            continue
        cell = img.crop((val_x0, yt, val_x1, yb))
        values[GI_FIELDS[i]] = _strip_noise(ocr(cell, psm=7,
                                              lang=template.get("OCR_LANG", OCR_LANG)))

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
    _render_clip(page, fitz.Rect(_px(template.get("PHOTO_BOX", PHOTO_BOX), w, h)),
                 img_dpi, photo_path)

    space_path = os.path.join(img_dir, f"space_p{page_idx + 1}.png")
    _render_clip(page, fitz.Rect(_px(template.get("SPACE_BOX", SPACE_BOX), w, h)),
                 img_dpi, space_path)

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


def _select_title_from_box(page: "fitz.Page", template: dict, page_text: str) -> str:
    """Pick a likely title from the TITLE_BOX words, fall back to page_text scan."""
    try:
        box = template.get("TITLE_BOX", TITLE_BOX)
        candidates = _words_in_box_lines(page, box)
        # Exclude common headers/labels and short numeric lines
        exclude = ("물류센터사진", "사진", "개요", "조감도", "Location",
                   "Site Plan", "Logistics")
        # Prefer a candidate that explicitly contains '물류센터'
        for ln in candidates:
            if not ln:
                continue
            if "물류센터" in ln and not any(ex in ln for ex in exclude):
                return clean_title(ln)
        for ln in candidates:
            if not ln:
                continue
            if any(ex in ln for ex in exclude):
                continue
            if re.search(r"[가-힣]", ln) and not re.match(r"^[\d\W]+$", ln):
                return clean_title(ln)
    except Exception:
        pass
    # fallback: look for the numbered headline in the page text
    m = re.search(r"^\s*\d+\.\s*(.+)$", page_text, re.M)
    return clean_title(m.group(1)) if m else ""


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

    # If box-based selection failed, fallback to heuristics over the page lines
    if not title:
        def _norm(s: str) -> str:
            return re.sub(r"\s+", "", s).strip().lower()

        # Prefer an explicit '물류센터' title or a numbered heading if box failed; fall back to
        # the first reasonably long Hangul-containing line near the top.
        for idx, ln in enumerate(lines):
            s = ln.lstrip("┃| ")
            if "물류센터" in s and not any(x in s for x in ("사진", "조감도", "Site Plan", "Location")):
                if "개요" in s or (s.startswith("물류센터") and (s.startswith("물류센터개요") or s.startswith("물류센터사진"))):
                    for look in range(idx + 1, min(idx + 4, len(lines))):
                        cand = lines[look]
                        if re.search(r"[가-힣]", cand) and len(cand) > 4 and not any(x in cand for x in ("사진", "조감도")):
                            title = clean_title(cand)
                            break
                    if title:
                        break
                    title = clean_title(s)
                    break
                else:
                    title = clean_title(s)
                    break
        if not title:
            for ln in lines:
                if re.match(r"^\d+\.\s*\S", ln):
                    title = clean_title(ln)
                    break
        if not title:
            for ln in lines[:6]:
                if re.search(r"[가-힣]", ln) and len(ln) > 6:
                    title = clean_title(ln)
                    break

    # Flexible label/value parsing: accept "라벨: 값" on one line or label
    # then value on the next line. Match labels fuzzily (ignore spaces/case).
    values: dict[str, str] = {}

    def _norm(s: str) -> str:
        return re.sub(r"\s+", "", s).strip().lower()

    def match_field(label: str) -> str | None:
        nl = _norm(label)
        # check explicit synonyms first
        for syn, canon in GI_SYNONYMS.items():
            if nl == _norm(syn):
                if canon in GI_FIELDS:
                    return canon
        for fld in GI_FIELDS:
            if nl == _norm(fld) or _norm(fld) in nl or nl in _norm(fld):
                return fld
        return None

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
    # For templates with a known SPACE_BOX, render that region; fallback to word-based rect
    try:
        if template.get("SPACE_BOX"):
            _render_clip(page, _rect_from_frac(template.get("SPACE_BOX"), page), crop_dpi, space_path)
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
            rec = extract_page(render_page(page, dpi), i, img_dir, template)
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
