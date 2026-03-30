"""투자심사보고서에서 데이터를 추출하는 모듈."""
import re
from dataclasses import dataclass, field
from docx import Document


@dataclass
class InvestmentReportData:
    # 회사 개요 (Table 1)
    company_name: str = ""
    representative: str = ""
    address: str = ""
    establishment_date: str = ""     # 설립일
    paid_in_capital: str = ""        # 납입자본금
    par_value: str = ""              # 액면가
    num_employees: str = ""          # 인력현황
    business_description: str = ""   # 주요사업
    business_registration: str = ""  # 사업자등록번호

    # 일정/담당 (Table 0)
    review_date: str = ""            # 투심 일자
    committee_date: str = ""         # 조합투심위
    contract_date: str = ""          # 계약일
    fund_date: str = ""              # 자금집행일
    fund_name: str = ""              # 투자 계정 (펀드명)
    discoverer: str = ""             # 발굴자 (기여율)
    reviewer: str = ""               # 심사자 (기여율)
    post_manager: str = ""           # 사후관리자 (기여율)

    # 투자 조건 (본문 텍스트)
    investment_amount: str = ""      # 투자금액
    issue_price: str = ""            # 투자단가
    total_shares: str = ""           # 인수주식수
    share_ratio: str = ""            # 지분율
    stock_type: str = ""             # 투자방식 (상환전환우선주 등)
    pre_value: str = ""              # Pre-Value 기업가치
    post_value: str = ""             # Post-Value 기업가치

    # 투자 조건 상세
    duration: str = ""               # 존속기간
    redemption_terms: str = ""       # 상환조건
    conversion_terms: str = ""       # 전환조건
    refixing_terms: str = ""         # Refixing
    other_terms: str = ""            # 기타
    fund_usage: str = ""             # 투자금 사용용도

    # 위약벌/이자율
    buyback_rate: str = ""
    penalty_rate: str = ""
    delay_rate: str = ""

    # 동반투자
    co_investors: list = field(default_factory=list)  # [(이름, 금액, 단가), ...]

    # 발굴경위
    discovery_background: str = ""

    # 경고
    warnings: list = field(default_factory=list)


def extract_report_data(filepath: str) -> InvestmentReportData:
    """투자심사보고서에서 데이터를 추출한다. DOCX와 PDF 모두 지원."""
    ext = filepath.lower().rsplit('.', 1)[-1] if '.' in filepath else ''
    data = InvestmentReportData()

    if ext == 'pdf':
        from extractors.pdf_extractor import extract_text_from_pdf
        full_text = extract_text_from_pdf(filepath)
        paragraphs = [p.strip() for p in full_text.split('\n') if p.strip()]
        doc = None
        tables = []

        # PDF에서는 텍스트 기반으로 모든 데이터 추출
        _extract_from_pdf_text(full_text, data)
    else:
        doc = Document(filepath)
        tables = doc.tables
        paragraphs = [p.text.strip() for p in doc.paragraphs]
        full_text = "\n".join(paragraphs)

        # Table 0: 일정/담당자
        if len(tables) > 0:
            _extract_table0(tables[0], data)

        # Table 1: 회사 개요
        if len(tables) > 1:
            _extract_table1(tables[1], data)

        # Table 3: 주주현황 → 지분율
        if len(tables) > 3:
            _extract_shareholder_table(tables, data)

    # 본문 텍스트에서 투자 조건 추출
    _extract_from_text(full_text, data, doc)

    return data


def _extract_from_pdf_text(full_text: str, data: InvestmentReportData):
    """PDF 텍스트에서 테이블 없이 모든 데이터를 텍스트 기반으로 추출."""
    # 회사명
    m = re.search(r'[㈜(주)]\s*(\S+)', full_text)
    if m:
        data.company_name = "㈜" + m.group(1)

    # 대표이사
    m = re.search(r'대표이사[:\s]*([가-힣]{2,4})', full_text)
    if m:
        data.representative = m.group(1)

    # 사업자등록번호
    m = re.search(r'(\d{3}-\d{2}-\d{5})', full_text)
    if m:
        data.business_registration = m.group(1)

    # 주소
    m = re.search(r'(서울|경기|인천|부산|대구|광주|대전|울산|세종|강원|충북|충남|전북|전남|경북|경남|제주)\S*\s*\S+\s*\S+[^)\n]{5,50}', full_text)
    if m:
        data.address = m.group(0).strip()

    # 설립일
    m = re.search(r'설립일[:\s]*([\d.]+)', full_text)
    if m:
        data.establishment_date = m.group(1)

    # 펀드명
    m = re.search(r'(20\d{2}\s*\S*\s*\S*\s*\S*\s*\d+호\s*펀드)', full_text)
    if m:
        data.fund_name = m.group(1)

    # 담당자
    m = re.search(r'발굴[:\s]*(\S+\s*\(\d+%\))', full_text)
    if m:
        data.discoverer = m.group(1)
    m = re.search(r'심사[:\s]*(\S+\s*\(\d+%\)(?:\s*,?\s*\S+\s*\(\d+%\))*)', full_text)
    if m:
        data.reviewer = m.group(1)
    m = re.search(r'사후관리[:\s]*(\S+\s*\(\d+%\))', full_text)
    if m:
        data.post_manager = m.group(1)


def _extract_table0(table, data: InvestmentReportData):
    """Table 0: 일정, 투자계정, 담당자 정보."""
    for row in table.rows:
        cells = [cell.text.strip() for cell in row.cells]
        if len(cells) < 2:
            continue

        row_text = " ".join(cells)

        if '투심 일자' in row_text or '투심일자' in row_text:
            data.review_date = cells[-1]
        elif '조합투심위' in row_text:
            data.committee_date = cells[-1]
        elif '계약일' in row_text:
            data.contract_date = cells[-1]
        elif '자금집행일' in row_text:
            data.fund_date = cells[-1]
        elif '투자 계정' in row_text or '투자계정' in row_text:
            data.fund_name = cells[-1]
        elif '발굴' in row_text and '심사' not in row_text:
            data.discoverer = cells[-1]
        elif '심사' in row_text:
            data.reviewer = cells[-1]
        elif '사후관리' in row_text:
            data.post_manager = cells[-1]


def _extract_table1(table, data: InvestmentReportData):
    """Table 1: 회사 개요."""
    for row in table.rows:
        cells = [cell.text.strip() for cell in row.cells]
        if len(cells) < 2:
            continue
        row_text = " ".join(cells)

        if '회사명' in row_text:
            for i, c in enumerate(cells):
                if '회사명' in c and i + 1 < len(cells):
                    data.company_name = cells[i + 1]
                    break
        if '대표이사' in row_text:
            for i, c in enumerate(cells):
                if '대표이사' in c and i + 1 < len(cells):
                    # 공백 제거 (강 귀 선 → 강귀선)
                    name = cells[i + 1].replace('\xa0', '').replace(' ', '').strip()
                    data.representative = name
                    break
        if '본사주소' in row_text or '본사 주소' in row_text:
            # 주소는 보통 병합 셀 → 마지막 비빈칸
            for c in reversed(cells):
                if c and '본사' not in c and len(c) > 5:
                    data.address = c.replace('\n', ' ').strip()
                    break
        if '납입자본금' in row_text:
            for i, c in enumerate(cells):
                if '납입자본금' in c and i + 1 < len(cells):
                    data.paid_in_capital = cells[i + 1]
                    break
        if '액면가' in row_text:
            for i, c in enumerate(cells):
                if '액면가' in c and i + 1 < len(cells):
                    data.par_value = cells[i + 1]
                    break
        if '설립일' in row_text:
            for i, c in enumerate(cells):
                if '설립일' in c and i + 1 < len(cells):
                    data.establishment_date = cells[i + 1]
                    break
        if '인력현황' in row_text:
            for i, c in enumerate(cells):
                if '인력현황' in c and i + 1 < len(cells):
                    data.num_employees = cells[i + 1]
                    break
        if '주요사업' in row_text:
            for c in reversed(cells):
                if c and '주요사업' not in c and len(c) > 3:
                    data.business_description = c
                    break


def _extract_shareholder_table(tables, data: InvestmentReportData):
    """주주현황 테이블에서 해당 펀드의 지분율을 추출."""
    for table in tables:
        # 합계행에서 총 주식수 추출
        total_post_shares = 0
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            if '합계' in " ".join(cells):
                for c in cells:
                    m = re.match(r'^([\d,]+)$', c.strip())
                    if m:
                        val = int(m.group(1).replace(',', ''))
                        if val > total_post_shares:
                            total_post_shares = val

        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            row_text = " ".join(cells)
            # 케이런벤처스 행 찾기
            if '케이런' in row_text and ('주' in row_text or '%' in row_text):
                # 주식수와 합계행에서 지분율 직접 계산 시도
                shares_in_row = []
                for c in cells:
                    m_shares = re.search(r'([\d,]+)', c)
                    if m_shares:
                        try:
                            val = int(m_shares.group(1).replace(',', ''))
                            if val > 100:  # 주식수로 보이는 값
                                shares_in_row.append(val)
                        except ValueError:
                            pass

                # 투자 후 지분율 (마지막에서 두번째 셀이 보통 투자 후 지분율)
                ratios = []
                for c in cells:
                    m = re.search(r'(\d+\.\d+)\s*%', c)
                    if m:
                        ratios.append(m.group(1))
                if ratios:
                    # 주식수와 합계가 있으면 직접 계산하여 더 정확한 값 사용
                    if total_post_shares > 0 and shares_in_row:
                        our_shares = max(shares_in_row)
                        calculated = round(our_shares / total_post_shares * 100, 2)
                        data.share_ratio = f"{calculated}%"
                        return
                    data.share_ratio = ratios[-1] + "%"
                    return
                # 소수점 없는 경우
                for c in cells:
                    m = re.search(r'(\d+)\s*%', c)
                    if m:
                        data.share_ratio = m.group(1) + "%"
                        return


def _extract_from_text(full_text: str, data: InvestmentReportData, doc=None):
    """본문 텍스트에서 투자조건, 동반투자 등 추출."""
    # 투자금액
    m = re.search(r'총\s*([\d,]+)\s*원.*?매수', full_text)
    if m:
        data.investment_amount = m.group(1) + "원"
    if not data.investment_amount:
        m = re.search(r'총\s*투자금액\s*([\d,]+)\s*원', full_text)
        if m:
            data.investment_amount = m.group(1) + "원"
    if not data.investment_amount:
        # "1,999,999,500원" 패턴
        m = re.search(r'([\d,]{7,})\s*원.*?규모', full_text)
        if m:
            data.investment_amount = m.group(1) + "원"

    # 투자단가 (주당)
    m = re.search(r'주당\s*([\d,]+)\s*원에', full_text)
    if m:
        data.issue_price = m.group(1) + "원"

    # 인수주식수
    m = re.search(r'([\d,]+)\s*주를?\s*주당', full_text)
    if m:
        data.total_shares = m.group(1) + "주"

    # 투자방식
    m = re.search(r'(상환전환우선주|전환우선주|상환우선주|보통주)\s*([\d,]+)\s*주', full_text)
    if m:
        data.stock_type = m.group(1)
        if not data.total_shares:
            data.total_shares = m.group(2) + "주"

    # Pre/Post 기업가치
    m = re.search(r'Pre.*?(\d+)\s*억', full_text)
    if m:
        data.pre_value = m.group(1) + "억원"
    m = re.search(r'Post.*?(\d+)\s*억', full_text)
    if m:
        data.post_value = m.group(1) + "억원"

    # 존속기간 / 상환 / 전환 / Refixing
    m = re.search(r'만기\s*(\d+)\s*년', full_text)
    if m:
        data.duration = m.group(1) + "년"

    m = re.search(r'발행\s*(\d+)\s*년\s*후.*?상환.*?YTM\s*(\d+)%', full_text)
    if m:
        data.redemption_terms = f"{m.group(1)}년후부터 상환청구 가능, 연복리 {m.group(2)}%"

    m = re.search(r'IPO/M&A\s*(\d+)%\s*리픽싱', full_text)
    if m:
        data.refixing_terms = f"IPO/M&A {m.group(1)}%"

    # 투자금 사용용도 - 테이블에서 우선 추출
    if doc:
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                if len(cells) >= 2 and '투자금 사용용도' in cells[0] and '위반' not in cells[0]:
                    data.fund_usage = cells[1].strip()
                    break
            if data.fund_usage:
                break
    # 텍스트에서 fallback
    if not data.fund_usage:
        for para in full_text.split('\n'):
            if '운영자금' in para or '설비투자' in para or '연구개발' in para:
                if len(para) < 200:
                    data.fund_usage = para.strip()
                    break

    # 동반투자
    m = re.search(r'동반투자기관[：:]?\s*(.+)', full_text)
    if m:
        co_text = m.group(1)
        # "산업은행 100억, 하나벤처스 30억" 등 파싱
        pairs = re.findall(r'(\S+)\s+(\d+)억', co_text)
        for name, amount in pairs:
            data.co_investors.append((name, f"{amount}억원", ""))

    # 사업자등록번호 - 테이블에서 우선 추출, 없으면 텍스트에서
    if not doc:
        return
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            row_text = " ".join(cells)
            if '사업자' in row_text or '등록번호' in row_text:
                for c in cells:
                    m = re.search(r'(\d{3}-\d{2}-\d{5})', c)
                    if m:
                        data.business_registration = m.group(1)
                        break
                if data.business_registration:
                    break
