"""
Streamlit UI for the MOVE PDF -> Excel extractor.

Drag-and-drop a property PDF, watch it get parsed, preview the result table and
the cropped images, then download the generated .xlsx.

Run:
    ./venv/bin/streamlit run app.py
"""
from __future__ import annotations

import io
import os
import tempfile
from datetime import date

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

from main import (COLUMNS, IMG_COLS, extract_records, merge_excel_bytes,
                  ocr_available, write_excel)
from templates import TEMPLATES, DEFAULT_TEMPLATE

st.set_page_config(page_title="물류센터 PDF → Excel 추출기",
                   page_icon="📄", layout="wide")

TEXT_COLUMNS = [c for c in COLUMNS if c not in IMG_COLS]
GALLERY_LIMIT = 24


def build_row(rec: dict, broker: str, info_date: str) -> dict:
    """Mirror the row defaults used by write_excel(), text columns only."""
    return {
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
        "시설특이사항": rec["시설특이사항"],
        "기타": "",
        "정보확인일자": info_date,
    }


def run_extraction(pdf_bytes: bytes, template_name: str, broker: str,
                   info_date: str, img_width: int, dpi: int,
                   img_dpi: int) -> dict | None:
    """Parse the PDF and return everything the UI needs (all in-memory).

    The PDF is read from memory and never written to disk; the cropped images
    and the intermediate workbook live in a TemporaryDirectory that is deleted
    automatically once their bytes have been read out below.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total = len(doc)
    bar = st.progress(0.0, text=f"0 / {total} 페이지 처리 중…")

    def on_progress(i, n, rec):
        label = rec["창고명"] if rec else "건너뜀"
        bar.progress((i + 1) / n, text=f"{i + 1} / {n} 페이지 — {label}")

    with tempfile.TemporaryDirectory() as tmp:  # auto-removed on block exit
        img_dir = os.path.join(tmp, "imgs")
        records, skipped = extract_records(doc, img_dir, dpi=dpi,
                           img_dpi=img_dpi,
                           progress=on_progress,
                           template_name=template)
        bar.empty()
        if not records:
            return None

        xlsx_path = os.path.join(tmp, "result.xlsx")
        write_excel(records, xlsx_path, broker, info_date, img_width)
        with open(xlsx_path, "rb") as fh:
            xlsx_bytes = fh.read()

        rows, gallery = [], []
        for idx, rec in enumerate(records):
            rows.append(build_row(rec, broker, info_date))
            if idx < GALLERY_LIMIT:
                with open(rec["_photo"], "rb") as fh:
                    photo = fh.read()
                with open(rec["_space"], "rb") as fh:
                    space = fh.read()
                gallery.append((rec["창고명"], photo, space))

    return {
        "xlsx": xlsx_bytes,
        "rows": rows,
        "gallery": gallery,
        "skipped": skipped,
        "count": len(records),
        "pages": total,
    }


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("📄 물류센터 PDF → Excel 추출기")
st.caption("PDF를 끌어다 놓으면 매물 페이지를 표로 추출하고 사진·공실현황 이미지를 "
           "넣은 엑셀 파일을 만들어 드립니다.")

if not ocr_available():
    st.warning("Tesseract OCR가 설치되어 있지 않습니다. 텍스트 레이어가 있는 PDF는 "
               "정상 동작하지만, 스캔(이미지) PDF는 추출되지 않습니다.  "
               "설치: `brew install tesseract tesseract-lang`")

with st.sidebar:
    st.header("옵션")
    template = st.selectbox("브로커 템플릿 선택", list(TEMPLATES.keys()),
                            index=list(TEMPLATES.keys()).index(DEFAULT_TEMPLATE),
                            help="PDF 레이아웃에 맞는 브로커 템플릿을 선택하세요.")
    broker = TEMPLATES[template].get("label", template)
    st.markdown(f"**중개인/임대인명:** {broker}")
    info_date = st.date_input("정보 확인일자", value=date.today()).strftime("%d/%m/%Y")
    img_width = st.slider("엑셀 이미지 너비(px)", 200, 600, 360, 20)
    with st.expander("고급 설정"):
        img_dpi = st.slider("이미지 렌더 DPI", 100, 300, 150, 10)
        ocr_dpi = st.slider("OCR 렌더 DPI (스캔 PDF)", 150, 400, 300, 10)
    # Sidebar: add quick access to Excel merge at the bottom
    st.markdown("---")
    st.header("📥 엑셀 파일 병합")
    merge_excels_sb = st.file_uploader(
        "병합할 Excel 파일들을 모두 업로드하세요.", type=["xlsx"],
        accept_multiple_files=True, key="merge_sidebar",
        help="여러 브로커별 엑셀 파일을 한 번에 선택하면 하나의 통합 워크북으로 병합합니다.")

    if merge_excels_sb:
        if len(merge_excels_sb) < 2:
            st.info("2개 이상의 Excel 파일을 업로드해야 병합할 수 있습니다.")
        else:
            if st.button("📦 엑셀 병합 (사이드바)", key="merge_sidebar_btn"):
                with st.spinner("엑셀 파일을 병합하는 중…"):
                    workbook_bytes = [f.getvalue() for f in merge_excels_sb]
                    try:
                        merged_bytes = merge_excel_bytes(workbook_bytes)
                        st.success(f"{len(merge_excels_sb)}개의 엑셀 파일을 성공적으로 병합했습니다.")
                        st.download_button(
                            "⬇️ 병합된 엑셀 파일 다운로드",
                            data=merged_bytes,
                            file_name="merged_workbook.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="merge_sidebar_dl",
                        )
                    except Exception as exc:
                        st.error(f"엑셀 병합에 실패했습니다: {exc}")

uploaded = st.file_uploader("PDF 파일을 여기에 끌어다 놓으세요", type=["pdf"],
                            accept_multiple_files=False)

if uploaded is not None:
    file_sig = (uploaded.name, uploaded.size, template, broker, info_date,
                img_width, img_dpi, ocr_dpi)
    if st.session_state.get("sig") != file_sig:
        with st.spinner(f"'{uploaded.name}' 분석 중…"):
            result = run_extraction(uploaded.getvalue(), template, broker,
                                    info_date, img_width, ocr_dpi, img_dpi)
        st.session_state["sig"] = file_sig
        st.session_state["result"] = result
        st.session_state["name"] = uploaded.name

result = st.session_state.get("result")

if uploaded is not None and result is None:
    st.error("이 PDF에서 매물 페이지를 찾지 못했습니다. 템플릿이 다른 PDF일 수 있습니다.")

elif result is not None:
    c1, c2, c3 = st.columns(3)
    c1.metric("추출된 매물 수", f"{result['count']} 건")
    c2.metric("건너뛴 페이지", f"{result['skipped']} 쪽")
    c3.metric("총 페이지", f"{result['pages']} 쪽")

    out_name = os.path.splitext(st.session_state.get("name", "result"))[0] + ".xlsx"
    st.download_button(
        "⬇️  엑셀 파일 다운로드", data=result["xlsx"], file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", width="stretch")

    st.subheader("추출 결과 (텍스트 항목)")
    df = pd.DataFrame(result["rows"], columns=TEXT_COLUMNS)
    df.index = range(1, len(df) + 1)
    st.dataframe(df, width="stretch", height=420)

    st.subheader(f"이미지 미리보기 (처음 {len(result['gallery'])}건)")
    for title, photo, space in result["gallery"]:
        st.markdown(f"**{title}**")
        g1, g2 = st.columns([1, 2])
        g1.image(photo, caption="사진", width="stretch")
        g2.image(space, caption="공실현황", width="stretch")
        st.divider()
else:
    st.info("좌측에서 옵션을 설정한 뒤 PDF를 업로드하세요.")

st.markdown("---")
st.header("📥 엑셀 파일 병합")
merge_excels = st.file_uploader(
    "병합할 Excel 파일들을 모두 업로드하세요.", type=["xlsx"],
    accept_multiple_files=True,
    help="여러 브로커별 엑셀 파일을 한 번에 선택하면 하나의 통합 워크북으로 병합합니다.")

if merge_excels:
    if len(merge_excels) < 2:
        st.info("2개 이상의 Excel 파일을 업로드해야 병합할 수 있습니다.")
    else:
        if st.button("📦 엑셀 병합"):
            with st.spinner("엑셀 파일을 병합하는 중…"):
                workbook_bytes = [uploaded.getvalue() for uploaded in merge_excels]
                try:
                    merged_bytes = merge_excel_bytes(workbook_bytes)
                    download_name = "merged_workbook.xlsx"
                    st.success(f"{len(merge_excels)}개의 엑셀 파일을 성공적으로 병합했습니다.")
                    st.download_button(
                        "⬇️ 병합된 엑셀 파일 다운로드",
                        data=merged_bytes,
                        file_name=download_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        width="stretch",
                    )
                except Exception as exc:
                    st.error(f"엑셀 병합에 실패했습니다: {exc}")
