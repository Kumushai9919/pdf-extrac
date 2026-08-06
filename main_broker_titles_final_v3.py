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
import posixpath
import re
import shutil
import sys
import zipfile
from collections import Counter
import xml.etree.ElementTree as ET
from pathlib import Path
from copy import deepcopy
from datetime import date
from difflib import SequenceMatcher
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
from PIL import Image, ImageOps
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
    "주소지": "소재지",
    "위치": "소재지",
    "소재": "소재지",
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
    m = re.search(r"([가-힣]+?(?:시|군|구))(?:([가-힣]+구))?", addr)
    if m:
        si = m.group(1)
        if m.group(2):
            si = f"{si} {m.group(2)}"
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
def _repair_cbre_mixed_latin_title(
    title_crop: Image.Image,
    title: str,
    template: dict,
) -> str:
    """Rebuild CBRE titles and repair low-confidence Latin codes such as JW."""
    if template.get("name") != "CBRE":
        return title

    try:
        data = pytesseract.image_to_data(
            title_crop,
            lang=template.get("OCR_LANG", OCR_LANG),
            config="--psm 7",
            output_type=pytesseract.Output.DICT,
        )
    except Exception:
        return title

    tokens: list[str] = []
    did_repair = False
    for index, raw in enumerate(data.get("text", [])):
        word = str(raw).strip()
        if not word:
            continue

        # Exclude the visual exclusivity badge from the warehouse name.
        if "전속" in word:
            continue

        try:
            confidence = float(data["conf"][index])
        except (TypeError, ValueError):
            confidence = -1

        repaired = word
        if (
            confidence < 60
            and re.search(r"[^가-힣A-Za-z]", word)
            and "물" not in word
        ):
            left = max(0, int(data["left"][index]) - 12)
            top = max(0, int(data["top"][index]) - 12)
            right = min(
                title_crop.width,
                int(data["left"][index]) + int(data["width"][index]) + 12,
            )
            bottom = min(
                title_crop.height,
                int(data["top"][index]) + int(data["height"][index]) + 12,
            )
            token_crop = title_crop.crop((left, top, right, bottom))
            try:
                latin = pytesseract.image_to_string(
                    token_crop,
                    lang="eng",
                    config="--psm 7",
                )
                latin = re.sub(r"[^A-Za-z0-9]", "", latin).upper()
                if re.fullmatch(r"[A-Z]{2,6}", latin):
                    repaired = latin
                    did_repair = True
            except Exception:
                pass

        tokens.append(repaired)

    if not did_repair:
        return _normalize_title_candidate(title, template)

    compact = re.sub(r"\s+", "", "".join(tokens))
    compact = re.sub(r"\[?전속\]?", "", compact)
    if "물류센터" not in compact:
        return title

    prefix, suffix = compact.split("물류센터", 1)
    prefix = re.sub(r"([가-힣])([A-Z])", r"\1 \2", prefix)
    prefix = re.sub(r"([A-Z])([가-힣])", r"\1 \2", prefix)
    # Keep short building codes attached to the location, e.g. 화성JW.
    prefix = re.sub(r"^([가-힣]+)\s+([A-Z]{2,4})$", r"\1\2", prefix)

    suffix = suffix.strip()
    suffix = re.sub(r"^(\d+)차부지", r"\1차 부지", suffix)
    suffix = re.sub(r"^(\d+)차", r"\1차 ", suffix)
    suffix = re.sub(r"\s+", " ", suffix).strip()

    rebuilt = f"{prefix.strip()} 물류센터"
    if suffix:
        rebuilt += f" {suffix}"

    rebuilt = _normalize_title_candidate(rebuilt, template)
    return rebuilt if _is_valid_title(rebuilt, template) else title


def _extract_title_from_image(img: Image.Image, template: dict) -> str:
    """OCR and validate only the configured title region of an image page."""
    w, h = img.size
    title_box = _px(template.get("TITLE_BOX", TITLE_BOX), w, h)
    crop = ImageOps.autocontrast(img.crop(title_box).convert("L"))

    candidates: list[str] = []
    psm_values = template.get("TITLE_OCR_PSMS", [7, 6, 11])
    for psm in dict.fromkeys(int(value) for value in psm_values):
        try:
            raw = ocr(
                crop,
                psm=psm,
                lang=template.get("OCR_LANG", OCR_LANG),
                scale=int(template.get("TITLE_OCR_SCALE", 2)),
            )
        except Exception:
            continue

        lines = [line.strip() for line in raw.splitlines() if line.strip()]
        candidates.extend(lines)
        if lines:
            candidates.append(" ".join(lines))

    title = _best_title_candidate(candidates, template)
    if title:
        title = _repair_cbre_mixed_latin_title(crop, title, template)
    return title


def _is_ocr_property_page(img: Image.Image, template: dict) -> bool:
    """Reject image-only cover/index/drawing pages before full extraction."""
    title = _extract_title_from_image(img, template)
    if not title:
        return False

    markers = template.get("OCR_PAGE_MARKERS", [])
    if not markers:
        return True

    w, h = img.size
    matched = 0
    for marker in markers:
        try:
            marker_box = _px(marker["box"], w, h)
            marker_img = ImageOps.autocontrast(
                img.crop(marker_box).convert("L")
            )
            marker_text = ocr(
                marker_img,
                psm=int(marker.get("psm", 11)),
                lang=template.get("OCR_LANG", OCR_LANG),
                scale=int(marker.get("scale", 2)),
            )
        except Exception:
            continue

        compact = re.sub(r"\s+", "", marker_text).lower()
        tokens = marker.get("tokens", [])
        if any(re.sub(r"\s+", "", token).lower() in compact for token in tokens):
            matched += 1

    required_min = int(template.get("OCR_PAGE_REQUIRED_MIN", len(markers)))
    if matched >= required_min:
        return True

    # Some valid image-only property pages only OCR one of the expected
    # marker regions, especially when the second heading is faint or broken.
    return matched >= 1


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
                    else:
                        parts = re.match(r"^([^:：\t]+?)\s+(.+)$", ln)
                        if parts:
                            fld = canonical_field(parts.group(1).strip())
                            if fld:
                                values.setdefault(fld, _strip_noise(parts.group(2).strip()))
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
    title = _extract_title_from_image(img, template)
    if template.get("name") == "Mateplus" and not title:
        title = _title_from_address(address, template)
    if not _is_valid_title(title, template):
        title = ""

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

    spec_field = template.get("SPEC_FIELD", "시설특이사항")

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

    output = {
        "창고명": title,
        "주소": address,
        "행정구역_도": do,
        "행정구역_시": si,
        "대지면적": clean_area(values.get("대지면적", "")),
        "연면적": clean_area(values.get("연면적", "")),
        "준공연도": year,
        "시설특이사항": "",
        "기타": "",
        "_photo": photo_path,
        "_space": space_path,
    }
    output[spec_field] = spec
    return output


# ---------------------------------------------------------------------------
# Text-layer extraction (for PDFs that already carry selectable text)
# ---------------------------------------------------------------------------
def is_property_page(text: str, template: dict) -> bool:
    """Detect a property sheet page using the selected template's keywords."""
    required_tokens = template.get("PAGE_REQUIRED_TOKENS", [])
    if required_tokens:
        text_compact = re.sub(r"\s+", "", text)
        required_count = sum(
            1
            for token in required_tokens
            if re.sub(r"\s+", "", token) in text_compact
        )
        required_min = int(
            template.get("PAGE_REQUIRED_MIN", len(required_tokens))
        )
        if required_count < required_min:
            return False

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


def _normalize_title_candidate(title: str, template: dict) -> str:
    """Normalize a title candidate without changing unrelated field values."""
    title = clean_title(title or "")
    if not title:
        return ""

    # Exclusivity badges are metadata, not part of the warehouse name.
    title = re.sub(r"\s*\[\s*전속[^\]]*\]\s*", " ", title)
    title = re.sub(r"\s+", " ", title).strip(" .,-")

    # Mapletree puts the nearest-IC distance on the same visual title line.
    # Every Mapletree warehouse title itself ends at '물류센터'.
    if template.get("name") == "Mapletree Korea" and "물류센터" in title:
        end = title.find("물류센터") + len("물류센터")
        title = title[:end].strip()

    if template.get("name") == "S1":
        # S1 places a small region badge above the title. OCR modes that treat
        # the crop as a block can join it to the actual warehouse name.
        title = re.sub(
            r"^\s*\[[^\]]*(?:권|전속)[^\]]*\]\s*",
            "",
            title,
        )

        # Red NEW labels are sometimes OCR'd as short Latin fragments after
        # the temperature suffix. Keep the title only through (상온/저온).
        temperature_title = re.match(
            r"^(.+?\((?:상온|저온)(?:\s*/\s*(?:상온|저온))?\))",
            title,
        )
        if temperature_title:
            title = temperature_title.group(1)

        title = re.sub(r"\s*/\s*", "/", title)
        title = re.sub(r"\s+", " ", title).strip()

    return clean_title(title)


def _is_valid_title(title: str, template: dict) -> bool:
    """Reject broker slogans, table values, and section headers."""
    title = _normalize_title_candidate(title, template)
    if not title:
        return False

    common_reject = (
        "%", "㎡", "용적률", "건폐율",
        "GENERAL INFORMATION", "SPACE AVAILABILITY",
        "PERSPECTIVE VIEW", "LOCATION", "LAYOUT",
        "Contact Point", "TABLE OF CONTENTS",
    )
    upper_title = title.upper()
    if any(str(token).upper() in upper_title for token in common_reject):
        return False
    # Reject area values such as "12,477평" without rejecting valid place
    # names such as "평택 물류센터".
    if re.search(r"\d[\d,.]*\s*(?:평|㎡)", title):
        return False

    reject_tokens = template.get("TITLE_REJECT_TOKENS", [])
    if any(str(token).upper() in upper_title for token in reject_tokens):
        return False

    compact = re.sub(r"\s+", "", title)
    if len(compact) < 4:
        return False
    if not re.search(r"[A-Za-z가-힣]", title):
        return False

    required = template.get("TITLE_REQUIRED_TOKENS", [])
    if required and not any(token in title for token in required):
        return False

    required_patterns = template.get("TITLE_REQUIRED_PATTERNS", [])
    if required_patterns and not any(
        re.search(pattern, title, flags=re.IGNORECASE)
        for pattern in required_patterns
    ):
        return False

    return True


def _title_score(title: str, template: dict) -> int:
    """Rank valid candidates while preserving all existing broker templates."""
    title = _normalize_title_candidate(title, template)
    if not _is_valid_title(title, template):
        return -1

    score = 0
    required = template.get("TITLE_REQUIRED_TOKENS", [])
    preferred = template.get("TITLE_PREFERRED_TOKENS", ["물류센터"])

    if required and any(token in title for token in required):
        score += 1000
    if any(token in title for token in preferred):
        score += 500

    # Do not let S1 badges like "[전속_ 경기중서권]" outrank the real title.
    if re.fullmatch(r"\s*\[[^\]]+\]\s*", title):
        score -= 300

    score += min(len(re.sub(r"\s+", "", title)), 100)
    return score


def _words_centered_in_box_lines(page: "fitz.Page", box: dict) -> list[str]:
    """Return visual text lines from words whose centres are inside a box."""
    rect = _rect_from_frac(box, page)
    selected: list[tuple[float, float, str]] = []

    for word in page.get_text("words"):
        x0, y0, x1, y1, value = word[:5]
        cx = (x0 + x1) / 2.0
        cy = (y0 + y1) / 2.0
        if rect.contains(fitz.Point(cx, cy)):
            selected.append((cy, x0, str(value)))

    selected.sort(key=lambda item: (item[0], item[1]))
    tolerance = max(2.0, page.rect.height * 0.004)
    rows: list[dict] = []

    for cy, x0, value in selected:
        if rows and abs(cy - rows[-1]["cy"]) <= tolerance:
            rows[-1]["items"].append((x0, value))
            rows[-1]["ys"].append(cy)
            rows[-1]["cy"] = sum(rows[-1]["ys"]) / len(rows[-1]["ys"])
        else:
            rows.append({
                "cy": cy,
                "ys": [cy],
                "items": [(x0, value)],
            })

    lines: list[str] = []
    for row in rows:
        parts = [
            value
            for _, value in sorted(row["items"], key=lambda item: item[0])
        ]
        line = clean_title(" ".join(parts))
        if line:
            lines.append(line)

    return lines


def _best_title_candidate(candidates: list[str], template: dict) -> str:
    """
    Select the most reliable title candidate.

    OCR is run with multiple PSM modes. A malformed PSM result can be longer
    than the correct title and previously won only because title length added
    points. Prefer candidates independently confirmed by multiple OCR modes,
    then apply the normal title-quality score.
    """
    normalized_candidates = [
        _normalize_title_candidate(candidate, template)
        for candidate in candidates
    ]
    normalized_candidates = [
        candidate
        for candidate in normalized_candidates
        if _is_valid_title(candidate, template)
    ]
    if not normalized_candidates:
        return ""

    frequencies = Counter(normalized_candidates)
    first_index: dict[str, int] = {}
    for index, candidate in enumerate(normalized_candidates):
        first_index.setdefault(candidate, index)

    scored: list[tuple[int, int, str]] = []
    for candidate, frequency in frequencies.items():
        score = _title_score(candidate, template)

        # Consensus between OCR modes is much more trustworthy than length.
        score += max(0, frequency - 1) * 250

        # Strong signs of fragmented OCR. Valid warehouse titles may contain
        # digits, but copyright symbols and several isolated tokens are noise.
        if re.search(r"[©®™]", candidate):
            score -= 1000
        if len(re.findall(r"(?<!\w)[A-Za-z0-9](?!\w)", candidate)) >= 3:
            score -= 400

        scored.append((score, first_index[candidate], candidate))

    scored.sort(key=lambda item: (-item[0], item[1]))
    return scored[0][2]


def _title_from_top_region(page: "fitz.Page", template: dict) -> str:
    """Safe text-layer fallback limited to the upper 15% of the page."""
    top_box = {"x0": 0.0, "x1": 1.0, "y0": 0.0, "y1": 0.15}
    return _best_title_candidate(
        _words_centered_in_box_lines(page, top_box),
        template,
    )


def _title_from_address(address: str, template: dict) -> str:
    """Lookup a MatePlus title from the cleaned address string."""
    if not address:
        return ""
    mapping = template.get("TITLE_BY_ADDRESS", {})
    if not isinstance(mapping, dict):
        return ""
    compact_address = re.sub(r"\s+", "", address)
    for key, title in mapping.items():
        if not key or not title:
            continue
        if re.sub(r"\s+", "", key) in compact_address:
            return clean_title(title)
    return ""


def _normalize_index_address(value: str) -> str:
    return re.sub(r"[^0-9A-Za-z가-힣]", "", value or "").lower()


def _unique_title(titles: set[str] | None) -> str:
    """Return the only title in a set; ambiguous or empty sets return blank."""
    if not titles or len(titles) != 1:
        return ""
    return next(iter(titles))


def _address_number_tokens(value: str) -> list[str]:
    """Keep Korean lot/building numbers such as 1549-1 as one token."""
    return re.findall(r"\d+(?:-\d+)?", value or "")


def _normalize_index_address(value: str) -> str:
    """
    Normalize harmless address spelling differences before matching.

    Examples:
        인천광역시 / 인천시 -> 인천
        경기도 / 경기 -> 경기
        whitespace and punctuation are ignored
    """
    normalized = value or ""

    administrative_aliases = (
        ("서울특별시", "서울"),
        ("부산광역시", "부산"),
        ("대구광역시", "대구"),
        ("인천광역시", "인천"),
        ("광주광역시", "광주"),
        ("대전광역시", "대전"),
        ("울산광역시", "울산"),
        ("세종특별자치시", "세종"),
        ("경기도", "경기"),
        ("강원특별자치도", "강원"),
        ("강원도", "강원"),
        ("충청북도", "충북"),
        ("충청남도", "충남"),
        ("전북특별자치도", "전북"),
        ("전라북도", "전북"),
        ("전라남도", "전남"),
        ("경상북도", "경북"),
        ("경상남도", "경남"),
        ("제주특별자치도", "제주"),
    )

    for long_name, short_name in administrative_aliases:
        normalized = normalized.replace(long_name, short_name)

    # Some PDFs duplicate one character during text extraction:
    # "인천광역시 시서구" -> "인천 시서구". Remove repeated admin suffixes
    # only when they occur as a standalone duplicate.
    normalized = re.sub(r"\b(시|도)\s+(?=[가-힣]+(?:시|군|구))", "", normalized)

    return re.sub(
        r"[^0-9A-Za-z가-힣]",
        "",
        normalized,
    ).lower()


def _match_mateplus_index_title(
    address: str,
    address_to_titles: dict[str, set[str]],
) -> str:
    """
    Return a MatePlus title only for a unique, confident address match.

    Safety rules:
        * exact unique match -> accept
        * unique containment match -> accept
        * unique fuzzy match with the same final lot/building number -> accept
        * ambiguous or weak match -> blank
        * never use page order
    """
    address_key = _normalize_index_address(address)
    if not address_key:
        return ""

    exact_title = _unique_title(address_to_titles.get(address_key))
    if exact_title:
        return exact_title

    containment_titles: set[str] = set()

    for saved_address, titles in address_to_titles.items():
        if not saved_address:
            continue
        if saved_address in address_key or address_key in saved_address:
            containment_titles.update(titles)

    containment_title = _unique_title(containment_titles)
    if containment_title:
        return containment_title
    if len(containment_titles) > 1:
        return ""

    # Compare numeric tokens after normalization so "394-20" and "39420"
    # are treated consistently on both sides.
    property_numbers = re.findall(r"\d+", address_key)
    final_property_number = (
        property_numbers[-1]
        if property_numbers
        else ""
    )

    if not final_property_number:
        return ""

    fuzzy_titles: set[str] = set()

    for saved_address, titles in address_to_titles.items():
        saved_numbers = re.findall(r"\d+", saved_address)
        final_saved_number = (
            saved_numbers[-1]
            if saved_numbers
            else ""
        )

        if final_saved_number != final_property_number:
            continue

        score = SequenceMatcher(
            None,
            address_key,
            saved_address,
        ).ratio()

        if score >= 0.82:
            fuzzy_titles.update(titles)

    return _unique_title(fuzzy_titles)


def _extract_mateplus_index(
    doc: "fitz.Document",
    template: dict,
) -> dict[str, set[str]]:
    """
    Read warehouse-title/address pairs from the MatePlus INDEX table.

    The visible word "INDEX" is vector artwork in this PDF and therefore isn't
    part of the selectable text layer. The page is identified using the real
    table headers instead.
    """
    index_page = None
    words: list[tuple] = []

    required_headers = {
        re.sub(r"\s+", "", token)
        for token in template.get(
            "INDEX_PAGE_REQUIRED_TOKENS",
            ["물류센터명", "소재지", "입주시기"],
        )
        if token and token != "INDEX"
    }

    if not required_headers:
        required_headers = {
            "물류센터명",
            "소재지",
            "입주시기",
        }

    for page in doc:
        candidate_words = page.get_text("words", sort=True)
        normalized_words = {
            re.sub(r"\s+", "", str(word[4]))
            for word in candidate_words
        }

        if required_headers.issubset(normalized_words):
            index_page = page
            words = candidate_words
            break

    if index_page is None:
        return {}

    title_header = next(
        (
            word
            for word in words
            if re.sub(r"\s+", "", str(word[4])) == "물류센터명"
        ),
        None,
    )
    address_header = next(
        (
            word
            for word in words
            if re.sub(r"\s+", "", str(word[4])) == "소재지"
        ),
        None,
    )
    next_header = next(
        (
            word
            for word in words
            if re.sub(r"\s+", "", str(word[4])) == "입주시기"
        ),
        None,
    )

    if title_header is None or address_header is None:
        return {}

    title_x0 = title_header[0] - 10
    title_x1 = address_header[0] - 5
    address_x0 = address_header[0] - 10
    address_x1 = (
        next_header[0] - 5
        if next_header is not None
        else index_page.rect.width * 0.56
    )
    header_bottom = max(title_header[3], address_header[3])

    tolerance = max(
        3.0,
        index_page.rect.height * 0.004,
    )
    selected: list[tuple[float, float, float, str]] = []

    for word in words:
        x0, y0, x1, y1, value = word[:5]
        center_y = (y0 + y1) / 2

        if center_y <= header_bottom:
            continue

        selected.append(
            (
                center_y,
                x0,
                x1,
                str(value).strip(),
            )
        )

    selected.sort(key=lambda item: (item[0], item[1]))
    rows: list[dict] = []

    for center_y, x0, x1, value in selected:
        if not value:
            continue

        if rows and abs(center_y - rows[-1]["center_y"]) <= tolerance:
            rows[-1]["items"].append((x0, x1, value))
            rows[-1]["ys"].append(center_y)
            rows[-1]["center_y"] = (
                sum(rows[-1]["ys"])
                / len(rows[-1]["ys"])
            )
        else:
            rows.append({
                "center_y": center_y,
                "ys": [center_y],
                "items": [(x0, x1, value)],
            })

    address_to_titles: dict[str, set[str]] = {}

    for row in rows:
        title_parts: list[str] = []
        address_parts: list[str] = []

        for x0, x1, value in sorted(
            row["items"],
            key=lambda item: item[0],
        ):
            center_x = (x0 + x1) / 2

            if title_x0 <= center_x < title_x1:
                title_parts.append(value)
            elif address_x0 <= center_x < address_x1:
                address_parts.append(value)

        title = clean_title(" ".join(title_parts))
        address = " ".join(address_parts).strip()

        if not title or not address:
            continue
        if title in {"물류센터명", "INDEX"}:
            continue

        address_key = _normalize_index_address(address)
        if not address_key:
            continue

        address_to_titles.setdefault(
            address_key,
            set(),
        ).add(title)

    return address_to_titles



def _extract_mateplus_title_from_blocks(
    page: "fitz.Page",
    template: dict,
) -> str:
    """
    Extract MatePlus title from PDF text blocks in the upper-left page area.
    Do not use Tesseract.
    """
    candidates: list[str] = []

    for block in page.get_text("blocks", sort=True):
        x0, y0, x1, y1, block_text = block[:5]

        cx = (x0 + x1) / 2
        cy = (y0 + y1) / 2

        # MatePlus title is in the upper-left part of the page.
        if cx > page.rect.width * 0.50:
            continue
        if cy > page.rect.height * 0.15:
            continue

        for line in str(block_text).splitlines():
            line = clean_title(line)

            if not line:
                continue

            compact = re.sub(r"\s+", "", line)

            if "물류센터" not in compact:
                continue

            if any(bad in line for bad in (
                "%",
                "㎡",
                "평",
                "GENERAL INFORMATION",
                "SPACE AVAILABILITY",
                "PERSPECTIVE VIEW",
            )):
                continue

            candidates.append(line)

    if not candidates:
        return ""

    candidates.sort(
        key=lambda value: len(re.sub(r"\s+", "", value)),
        reverse=True,
    )

    return candidates[0]



def _extract_title_with_ocr(
    page: "fitz.Page",
    template: dict,
) -> str:
    """
    OCR only the configured title box.

    This is a fallback for PDFs where the visible warehouse title is drawn as
    outlines/image content and therefore isn't present in the selectable text
    layer. It is intentionally limited to TITLE_BOX so it cannot accidentally
    pick values from GENERAL INFORMATION.
    """
    try:
        box = template.get("TITLE_BOX", TITLE_BOX)
        clip = _rect_from_frac(box, page)
        dpi = int(template.get("TITLE_OCR_DPI", 300))
        pix = page.get_pixmap(
            dpi=dpi,
            clip=clip,
            alpha=False,
        )
        image = Image.frombytes(
            "RGB",
            (pix.width, pix.height),
            pix.samples,
        )
    except Exception:
        return ""

    # The MatePlus title is dark text on white. Autocontrast improves OCR
    # without changing the source PDF or the exported Excel images.
    grayscale = ImageOps.autocontrast(image.convert("L"))
    candidates: list[str] = []
    psm_values = [
        int(template.get("TITLE_OCR_PSM", 7)),
        6,
        11,
    ]

    for psm in dict.fromkeys(psm_values):
        try:
            raw = ocr(
                grayscale,
                psm=psm,
                lang=template.get("OCR_LANG", OCR_LANG),
                scale=2,
            )
        except Exception:
            continue

        lines = [
            clean_title(line)
            for line in raw.splitlines()
            if line.strip()
        ]
        candidates.extend(lines)

        combined = clean_title(" ".join(lines))
        if combined:
            candidates.append(combined)

    return _best_title_candidate(candidates, template)


def _select_title_from_box(page: "fitz.Page", template: dict,
                           page_text: str) -> str:
    """Read the warehouse title from the configured top title region."""
    if template.get("TITLE_EXTRACT_MODE") == "mateplus_blocks":
        title = _extract_mateplus_title_from_blocks(page, template)
        if title:
            return title

    try:
        box = template.get("TITLE_BOX", TITLE_BOX)
        title = _best_title_candidate(
            _words_centered_in_box_lines(page, box),
            template,
        )
        if title:
            return title
    except Exception:
        pass

    # Some brokers render the title as vector outlines/image content.
    # OCR only the configured title box when the template explicitly allows it.
    if template.get("TITLE_OCR_FALLBACK", template.get("name") == "Mateplus"):
        title = _extract_title_with_ocr(page, template)
        if title:
            return title

    return _title_from_top_region(page, template)


def extract_text_page(page: "fitz.Page", page_idx: int,
                      img_dir: str, crop_dpi: int, template: dict) -> dict:
    """Parse one property sheet from its text layer + render the two images.

    Uses the template `TITLE_BOX` region to pick a reliable title when
    available, otherwise falls back to heuristics over the page text.
    """
    text = page.get_text()
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    # MatePlus property-page titles are vector artwork, not selectable text.
    # For TITLE_SOURCE="index", skip OCR and assign 창고명 later by matching
    # this page's extracted address to the selectable INDEX table.
    if (
        template.get("name") == "Mateplus"
        and template.get("TITLE_SOURCE") == "index"
    ):
        title = ""
    else:
        title = _select_title_from_box(page, template, text)

    # MatePlus fallback: scan the selectable top text lines for a visible 물류센터 title.
    if template.get("name") == "Mateplus" and not title:
        top_lines = [
            clean_title(line)
            for line in page.get_text("text", sort=True).splitlines()[:20]
            if line.strip()
        ]

        for candidate in top_lines:
            compact = re.sub(r"\s+", "", candidate)

            if "물류센터" not in compact:
                continue

            if any(bad in candidate for bad in (
                "%",
                "㎡",
                "평",
                "GENERAL INFORMATION",
                "SPACE AVAILABILITY",
                "PERSPECTIVE VIEW",
                "LOCATION",
            )):
                continue

            title = candidate
            break

    # TITLE_BOX selection already includes a safe top-of-page fallback.
    # Never scan GENERAL INFORMATION or arbitrary page lines for a title.
    if not _is_valid_title(title, template):
        title = ""

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

    address = clean_address(values.get("소재지", ""))
    if template.get("name") == "Mateplus" and not title:
        title = _title_from_address(address, template)

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
    if template.get("name") == "Mateplus" and not title:
        title = _title_from_address(address, template)
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
# Excel image optimization
# ---------------------------------------------------------------------------
# The building photo is only a visual reference, so it can be compressed
# aggressively. The availability table contains text, so it keeps more detail.
PHOTO_EXCEL_MAX_WIDTH = 700
PHOTO_EXCEL_JPEG_QUALITY = 65
SPACE_EXCEL_MAX_WIDTH = 1000
SPACE_EXCEL_COLORS = 128


def _flatten_to_rgb(image: Image.Image) -> Image.Image:
    """Convert an image to RGB while replacing transparency with white."""
    if image.mode in ("RGBA", "LA") or (
        image.mode == "P" and "transparency" in image.info
    ):
        rgba = image.convert("RGBA")
        background = Image.new("RGB", rgba.size, "white")
        background.paste(rgba, mask=rgba.getchannel("A"))
        return background

    return image.convert("RGB")


def _resize_to_max_width(image: Image.Image, max_width: int) -> Image.Image:
    """Downsample without enlarging an already small image."""
    if image.width <= max_width:
        return image.copy()

    new_height = max(
        1,
        round(image.height * max_width / image.width),
    )
    return image.resize(
        (max_width, new_height),
        Image.Resampling.LANCZOS,
    )


def _optimize_image_bytes(
    image_bytes: bytes,
    image_kind: str,
    original_filename: str = "image.png",
) -> tuple[bytes, str]:
    """
    Reduce XLSX image payload size without changing extraction/OCR.

    사진:
        JPEG, max 700 px wide, quality 65.
    공실현황:
        indexed PNG, max 1000 px wide, 128 colors for readable text.
    """
    with Image.open(BytesIO(image_bytes)) as source:
        source.load()

        if image_kind == "사진":
            # Avoid repeatedly recompressing an image that is already in the
            # optimized format and dimensions.
            if (
                source.format == "JPEG"
                and source.width <= PHOTO_EXCEL_MAX_WIDTH
            ):
                return image_bytes, Path(original_filename).with_suffix(
                    ".jpg"
                ).name

            image = _resize_to_max_width(
                _flatten_to_rgb(source),
                PHOTO_EXCEL_MAX_WIDTH,
            )
            output = BytesIO()
            image.save(
                output,
                format="JPEG",
                quality=PHOTO_EXCEL_JPEG_QUALITY,
                optimize=True,
                progressive=True,
                subsampling=2,
            )
            return (
                output.getvalue(),
                Path(original_filename).with_suffix(".jpg").name,
            )

        # 공실현황 is mainly a table/text image. Palette PNG keeps lines and
        # text sharper than low-quality JPEG while still reducing file size.
        if (
            source.format == "PNG"
            and source.mode == "P"
            and source.width <= SPACE_EXCEL_MAX_WIDTH
        ):
            return image_bytes, Path(original_filename).with_suffix(
                ".png"
            ).name

        image = _resize_to_max_width(
            _flatten_to_rgb(source),
            SPACE_EXCEL_MAX_WIDTH,
        )
        image = image.quantize(
            colors=SPACE_EXCEL_COLORS,
            method=Image.Quantize.MEDIANCUT,
            dither=Image.Dither.NONE,
        )
        output = BytesIO()
        image.save(
            output,
            format="PNG",
            optimize=True,
            compress_level=9,
        )
        return (
            output.getvalue(),
            Path(original_filename).with_suffix(".png").name,
        )


def _prepare_excel_image_file(path: str, image_kind: str) -> str:
    """Create an optimized sibling image and return its path."""
    source_path = Path(path)
    optimized_bytes, optimized_filename = _optimize_image_bytes(
        source_path.read_bytes(),
        image_kind,
        source_path.name,
    )

    suffix = Path(optimized_filename).suffix
    optimized_path = source_path.with_name(
        f"{source_path.stem}_excel{suffix}"
    )
    optimized_path.write_bytes(optimized_bytes)
    return str(optimized_path)


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
            "시설특이사항": rec.get("시설특이사항", ""),
            "기타": rec.get("기타", ""),
            "정보확인일자": info_date,
            "사진": "",
        }
        max_row_px = 24

        for name, value in row_values.items():
            ws.write(row_idx, col_index[name] - 1, value, cell_format)

        for name, path in (("공실현황", rec["_space"]), ("사진", rec["_photo"])):
            if not path:
                continue

            optimized_path = _prepare_excel_image_file(path, name)

            with Image.open(optimized_path) as im:
                ow, oh = im.size
            if ow == 0:
                continue

            scale = img_target_w / ow
            scaled_height = int(oh * scale)
            max_row_px = max(max_row_px, scaled_height + 10)

            ws.embed_image(
                row_idx,
                col_index[name] - 1,
                optimized_path,
                {
                    "description": name,
                    "x_scale": scale,
                    "y_scale": scale,
                },
            )

        ws.set_row_pixels(row_idx, max_row_px)

    ws.freeze_panes(1, 0)
    workbook.close()


# OOXML namespaces used by XlsxWriter's native "Place in Cell" images.
_XML_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_XML_DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_XML_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_XML_RICH = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
_XML_RICH_REL = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"


def _zip_part_name(zf: zipfile.ZipFile, name: str) -> str | None:
    """Return the actual ZIP member name, tolerating an optional leading slash."""
    names = set(zf.namelist())
    if name in names:
        return name

    alternate = "/" + name.lstrip("/")
    if alternate in names:
        return alternate

    return None


def _read_zip_xml(zf: zipfile.ZipFile, name: str) -> ET.Element:
    actual_name = _zip_part_name(zf, name)
    if actual_name is None:
        raise KeyError(name)
    return ET.fromstring(zf.read(actual_name))


def _resolve_ooxml_target(base_part: str, target: str) -> str:
    """Resolve an OOXML relationship target to a normalized ZIP part path."""
    if target.startswith("/"):
        return target.lstrip("/")

    return posixpath.normpath(
        posixpath.join(posixpath.dirname(base_part), target)
    )


def _first_worksheet_part(zf: zipfile.ZipFile) -> str:
    """Resolve the first worksheet XML part from workbook relationships."""
    workbook_root = _read_zip_xml(zf, "xl/workbook.xml")
    sheet = workbook_root.find(f".//{{{_XML_MAIN}}}sheet")

    if sheet is None:
        raise ValueError("The workbook has no worksheet.")

    relationship_id = sheet.attrib.get(f"{{{_XML_DOC_REL}}}id")
    relationships_root = _read_zip_xml(
        zf,
        "xl/_rels/workbook.xml.rels",
    )

    for relationship in relationships_root.findall(
        f"{{{_XML_PKG_REL}}}Relationship"
    ):
        if relationship.attrib.get("Id") == relationship_id:
            return _resolve_ooxml_target(
                "xl/workbook.xml",
                relationship.attrib["Target"],
            )

    raise ValueError("Could not resolve the first worksheet XML part.")


def extract_in_cell_images(
    workbook_data: bytes,
) -> dict[str, tuple[bytes, str, str]]:
    """
    Extract XlsxWriter/Excel native "Place in Cell" images from an XLSX.

    Returns:
        {
            "K2": (image_bytes, "image1.png", "공실현황"),
            "O2": (image_bytes, "image2.png", "사진"),
        }

    These images aren't exposed through openpyxl's ``worksheet._images``
    because they are stored as Excel rich-data values, not drawing objects.
    """
    images: dict[str, tuple[bytes, str, str]] = {}

    with zipfile.ZipFile(BytesIO(workbook_data)) as zf:
        required_parts = (
            "xl/metadata.xml",
            "xl/richData/rdrichvalue.xml",
            "xl/richData/richValueRel.xml",
            "xl/richData/_rels/richValueRel.xml.rels",
        )

        if any(_zip_part_name(zf, part) is None for part in required_parts):
            return images

        worksheet_part = _first_worksheet_part(zf)
        worksheet_root = _read_zip_xml(zf, worksheet_part)

        # Cell vm is a 1-based index into valueMetadata.
        image_cells: list[tuple[str, int]] = []

        for cell in worksheet_root.findall(f".//{{{_XML_MAIN}}}c"):
            cell_reference = cell.attrib.get("r")
            metadata_index = cell.attrib.get("vm")

            if not cell_reference or not metadata_index:
                continue

            try:
                image_cells.append(
                    (cell_reference, int(metadata_index) - 1)
                )
            except ValueError:
                continue

        if not image_cells:
            return images

        metadata_root = _read_zip_xml(zf, "xl/metadata.xml")
        value_metadata = metadata_root.findall(
            f"./{{{_XML_MAIN}}}valueMetadata/"
            f"{{{_XML_MAIN}}}bk"
        )
        future_metadata = metadata_root.findall(
            f"./{{{_XML_MAIN}}}futureMetadata/"
            f"{{{_XML_MAIN}}}bk"
        )

        rich_values_root = _read_zip_xml(
            zf,
            "xl/richData/rdrichvalue.xml",
        )
        rich_values = rich_values_root.findall(
            f"{{{_XML_RICH}}}rv"
        )

        rich_rel_root = _read_zip_xml(
            zf,
            "xl/richData/richValueRel.xml",
        )
        rich_relations = rich_rel_root.findall(
            f"{{{_XML_RICH_REL}}}rel"
        )

        package_rel_root = _read_zip_xml(
            zf,
            "xl/richData/_rels/richValueRel.xml.rels",
        )
        relationship_targets = {
            relationship.attrib["Id"]: relationship.attrib["Target"]
            for relationship in package_rel_root.findall(
                f"{{{_XML_PKG_REL}}}Relationship"
            )
        }

        rich_relation_part = "xl/richData/richValueRel.xml"

        for cell_reference, metadata_index in image_cells:
            try:
                # valueMetadata -> futureMetadata -> rich-value index.
                rc = value_metadata[metadata_index].find(
                    f"{{{_XML_MAIN}}}rc"
                )
                if rc is None:
                    continue

                future_index = int(rc.attrib["v"])
                rich_value_binding = future_metadata[future_index].find(
                    f".//{{{_XML_RICH}}}rvb"
                )
                if rich_value_binding is None:
                    continue

                rich_value_index = int(
                    rich_value_binding.attrib["i"]
                )
                rich_value = rich_values[rich_value_index]
                rich_value_fields = rich_value.findall(
                    f"{{{_XML_RICH}}}v"
                )

                if not rich_value_fields:
                    continue

                # First field is an index into richValueRel.xml.
                relation_index = int(rich_value_fields[0].text or "0")
                description = (
                    rich_value_fields[2].text
                    if len(rich_value_fields) > 2
                    and rich_value_fields[2].text
                    else cell_reference
                )

                relation_id = rich_relations[relation_index].attrib[
                    f"{{{_XML_DOC_REL}}}id"
                ]
                target = relationship_targets[relation_id]
                media_part = _resolve_ooxml_target(
                    rich_relation_part,
                    target,
                )
                actual_media_part = _zip_part_name(zf, media_part)

                if actual_media_part is None:
                    continue

                images[cell_reference] = (
                    zf.read(actual_media_part),
                    posixpath.basename(media_part),
                    description,
                )

            except (
                IndexError,
                KeyError,
                TypeError,
                ValueError,
                AttributeError,
            ):
                # A malformed rich-data entry shouldn't prevent the rest of
                # the workbook from being merged.
                continue

    return images


def _row_height_pixels(source_ws, row_number: int, has_image: bool) -> int:
    """Convert an openpyxl row height in points to XlsxWriter pixels."""
    row_dimension = source_ws.row_dimensions.get(row_number)

    if row_dimension is not None and row_dimension.height:
        return max(20, int(round(float(row_dimension.height) * 96 / 72)))

    return 200 if has_image else 24


def merge_excel_bytes(workbook_bytes: list[bytes]) -> bytes:
    """
    Merge extractor workbooks while preserving native in-cell images.

    Source workbooks are read with openpyxl for ordinary cell values only.
    Native images are extracted from the XLSX rich-data XML and embedded again
    into a newly generated XlsxWriter workbook.
    """
    if not workbook_bytes:
        raise ValueError("No workbooks provided for merging.")

    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("Warehouses")

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

    for column_zero_based, name in enumerate(COLUMNS):
        worksheet.write(
            0,
            column_zero_based,
            name,
            header_format,
        )

    worksheet.set_row_pixels(0, 24)

    column_index = {
        name: index
        for index, name in enumerate(COLUMNS)
    }
    image_column_indexes = {
        column_index["공실현황"],
        column_index["사진"],
    }

    default_widths = {
        "사용가능여부": 12,
        "중개인/임대인명": 14,
        "창고명": 26,
        "주소": 26,
        "행정구역_도": 14,
        "행정구역_시": 12,
        "대지면적": 20,
        "연면적": 20,
        "건축면적": 14,
        "준공연도": 12,
        "공실현황": 51.43,
        "시설특이사항": 38,
        "기타": 12,
        "정보확인일자": 14,
        "사진": 51.43,
    }
    maximum_widths = dict(default_widths)

    # XlsxWriter needs these streams to remain alive until workbook.close().
    image_streams: list[BytesIO] = []
    destination_row_zero_based = 1

    try:
        for source_data in workbook_bytes:
            source_workbook = load_workbook(
                filename=BytesIO(source_data),
                data_only=False,
                read_only=False,
            )
            source_ws = source_workbook.active

            source_header = [
                source_ws.cell(row=1, column=column).value
                for column in range(1, len(COLUMNS) + 1)
            ]

            if source_header != COLUMNS:
                raise ValueError(
                    "Cannot merge workbooks with incompatible headers."
                )

            source_images = extract_in_cell_images(source_data)

            # Preserve the widest source column for each field.
            for column_one_based, name in enumerate(COLUMNS, start=1):
                column_letter = get_column_letter(column_one_based)
                dimension = source_ws.column_dimensions.get(column_letter)

                if dimension is not None and dimension.width:
                    maximum_widths[name] = max(
                        maximum_widths.get(name, 0),
                        float(dimension.width),
                    )

            for source_row in range(2, source_ws.max_row + 1):
                image_refs = {
                    column_zero_based: (
                        f"{get_column_letter(column_zero_based + 1)}"
                        f"{source_row}"
                    )
                    for column_zero_based in image_column_indexes
                }

                has_image = any(
                    reference in source_images
                    for reference in image_refs.values()
                )

                normal_values = [
                    source_ws.cell(
                        row=source_row,
                        column=column_zero_based + 1,
                    ).value
                    for column_zero_based in range(len(COLUMNS))
                    if column_zero_based not in image_column_indexes
                ]

                # Ignore completely empty source rows.
                if not has_image and not any(
                    value not in (None, "")
                    for value in normal_values
                ):
                    continue

                for column_zero_based, name in enumerate(COLUMNS):
                    if column_zero_based in image_column_indexes:
                        continue

                    value = source_ws.cell(
                        row=source_row,
                        column=column_zero_based + 1,
                    ).value

                    worksheet.write(
                        destination_row_zero_based,
                        column_zero_based,
                        value,
                        cell_format,
                    )

                for column_zero_based in image_column_indexes:
                    source_reference = image_refs[column_zero_based]
                    image_info = source_images.get(source_reference)

                    if image_info is None:
                        # Never copy the source #VALUE! placeholder.
                        worksheet.write_blank(
                            destination_row_zero_based,
                            column_zero_based,
                            None,
                            cell_format,
                        )
                        continue

                    image_bytes, filename, description = image_info
                    image_kind = COLUMNS[column_zero_based]
                    image_bytes, filename = _optimize_image_bytes(
                        image_bytes,
                        image_kind,
                        filename,
                    )
                    image_stream = BytesIO(image_bytes)
                    image_streams.append(image_stream)

                    result = worksheet.embed_image(
                        destination_row_zero_based,
                        column_zero_based,
                        filename,
                        {
                            "image_data": image_stream,
                            "description": description,
                            "cell_format": cell_format,
                        },
                    )

                    if result != 0:
                        raise RuntimeError(
                            "Could not embed image from "
                            f"{source_reference}."
                        )

                worksheet.set_row_pixels(
                    destination_row_zero_based,
                    _row_height_pixels(
                        source_ws,
                        source_row,
                        has_image,
                    ),
                )
                destination_row_zero_based += 1

            source_workbook.close()

        for column_zero_based, name in enumerate(COLUMNS):
            worksheet.set_column(
                column_zero_based,
                column_zero_based,
                maximum_widths[name],
            )

        worksheet.freeze_panes(1, 0)
        workbook.close()

    except Exception:
        # Avoid returning a partially written workbook.
        try:
            workbook.close()
        except Exception:
            pass
        raise

    output.seek(0)
    return output.read()


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

    mateplus_address_titles: dict[str, set[str]] = {}

    if (
        template.get("name") == "Mateplus"
        and template.get("TITLE_SOURCE") == "index"
    ):
        mateplus_address_titles = _extract_mateplus_index(
            doc,
            template,
        )

    records: list[dict] = []
    skipped = 0
    total = len(doc)
    for i, page in enumerate(doc):
        text = page.get_text()
        rec = None
        if is_property_page(text, template):
            rec = extract_text_page(page, i, img_dir, img_dpi, template)
        elif (
            template.get("name") != "Mateplus"
            and len(text.strip()) < 50
            and ocr_ready
        ):
            rendered_page = render_page(page, dpi)
            if _is_ocr_property_page(rendered_page, template):
                rec = extract_page(
                    rendered_page,
                    i,
                    img_dir,
                    template,
                    crop_dpi=img_dpi,
                )
            else:
                skipped += 1
        else:
            # MatePlus property pages already contain selectable text.
            # Never OCR its cover/SITE PLAN pages, because Streamlit Cloud has
            # Tesseract installed and can otherwise create false records.
            skipped += 1
        if (
            rec is not None
            and template.get("name") == "Mateplus"
            and template.get("TITLE_SOURCE") == "index"
        ):
            # Strict business-data rule:
            # unique confident address match -> title;
            # otherwise leave blank. Never guess from page order.
            rec["창고명"] = _match_mateplus_index_title(
                rec.get("주소", ""),
                mateplus_address_titles,
            )

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