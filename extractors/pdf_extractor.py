"""PDF 파일에서 텍스트를 추출하는 모듈. 텍스트 PDF와 스캔 PDF 모두 지원."""
import os
import re
import fitz  # pymupdf
import pytesseract
from PIL import Image
import io


def extract_text_from_pdf(filepath: str) -> str:
    """PDF에서 텍스트를 추출한다. 텍스트 PDF는 직접 추출, 스캔본은 OCR."""
    doc = fitz.open(filepath)
    all_text = []
    ocr_used = False

    for page_num in range(len(doc)):
        page = doc[page_num]

        # 먼저 텍스트 직접 추출 시도
        text = page.get_text("text").strip()

        if len(text) < 50:
            # 텍스트가 거의 없으면 스캔본으로 판단 → OCR
            if not ocr_used:
                print(f"  [OCR] 스캔본 감지. OCR 처리 중...")
                ocr_used = True

            text = _ocr_page(page, page_num)

        all_text.append(text)

    doc.close()

    full_text = "\n\n".join(all_text)

    if ocr_used:
        print(f"  [OCR] {len(doc)} 페이지 OCR 완료 ({len(full_text)} 글자)")

    return full_text


def extract_tables_from_pdf(filepath: str) -> list:
    """PDF에서 테이블 형태의 데이터를 추출한다.
    반환: [{"rows": [["cell1", "cell2", ...], ...]}]
    """
    # PDF 테이블 추출은 텍스트 기반으로 간접 파싱
    text = extract_text_from_pdf(filepath)
    tables = _parse_tables_from_text(text)
    return tables


def _ocr_page(page, page_num: int, dpi: int = 300) -> str:
    """페이지를 이미지로 렌더링한 후 OCR로 텍스트 추출."""
    # 고해상도 이미지로 렌더링
    zoom = dpi / 72
    matrix = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=matrix)

    # PIL Image로 변환
    img_data = pix.tobytes("png")
    img = Image.open(io.BytesIO(img_data))

    # Tesseract OCR (한국어+영어)
    text = pytesseract.image_to_string(img, lang='kor+eng', config='--psm 6')

    return text.strip()


def _parse_tables_from_text(text: str) -> list:
    """OCR 텍스트에서 테이블 구조를 파싱한다.
    테이블은 탭/다중공백으로 구분된 행으로 인식."""
    tables = []
    current_table = []
    lines = text.split('\n')

    for line in lines:
        line = line.strip()
        if not line:
            if current_table and len(current_table) >= 2:
                tables.append({"rows": current_table})
            current_table = []
            continue

        # 탭이나 다중 공백으로 구분된 셀 감지
        cells = re.split(r'\t|  +|\|', line)
        cells = [c.strip() for c in cells if c.strip()]

        if len(cells) >= 2:
            current_table.append(cells)
        elif current_table:
            # 단일 셀 행이 테이블 중간에 오면 포함
            current_table.append([line])

    if current_table and len(current_table) >= 2:
        tables.append({"rows": current_table})

    return tables
