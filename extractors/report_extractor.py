"""투자심사보고서에서 데이터를 추출하는 모듈.
테이블 위치/구조에 무관하게 키워드 기반으로 자동 탐색한다."""
import re
from dataclasses import dataclass, field
from docx import Document


@dataclass
class InvestmentReportData:
    company_name: str = ""
    representative: str = ""
    address: str = ""
    establishment_date: str = ""
    paid_in_capital: str = ""
    par_value: str = ""
    num_employees: str = ""
    business_description: str = ""
    business_registration: str = ""

    review_date: str = ""
    committee_date: str = ""
    contract_date: str = ""
    fund_date: str = ""
    fund_name: str = ""
    discoverer: str = ""
    reviewer: str = ""
    post_manager: str = ""

    investment_amount: str = ""
    issue_price: str = ""
    total_shares: str = ""
    share_ratio: str = ""
    stock_type: str = ""
    pre_value: str = ""
    post_value: str = ""

    duration: str = ""
    redemption_terms: str = ""
    conversion_terms: str = ""
    refixing_terms: str = ""
    other_terms: str = ""
    fund_usage: str = ""

    buyback_rate: str = ""
    penalty_rate: str = ""
    delay_rate: str = ""

    co_investors: list = field(default_factory=list)
    discovery_background: str = ""
    warnings: list = field(default_factory=list)


# ── 지역 리스트 ──
_REGIONS = (
    '서울|경기|인천|부산|대구|광주|대전|울산|세종|강원|충북|충남'
    '|전북|전남|경북|경남|제주|서울특별시|대전광역시|대전시|부산광역시'
)


def extract_report_data(filepath: str) -> InvestmentReportData:
    ext = filepath.lower().rsplit('.', 1)[-1] if '.' in filepath else ''
    data = InvestmentReportData()

    if ext == 'pdf':
        from extractors.pdf_extractor import extract_text_from_pdf
        full_text = extract_text_from_pdf(filepath)
        doc = None
        _extract_from_pdf_text(full_text, data)
    else:
        doc = Document(filepath)
        paragraphs = [p.text.strip() for p in doc.paragraphs]
        full_text = "\n".join(paragraphs)
        _scan_all_tables(doc.tables, data)

    _extract_from_text(full_text, data, doc)
    return data


# ━━━━━━━━━━━━━━━ 테이블 자동 탐색 ━━━━━━━━━━━━━━━

def _scan_all_tables(tables, data: InvestmentReportData):
    """모든 테이블을 순회하며 키워드로 테이블 유형을 감지한다."""
    for table in tables:
        if not table.rows:
            continue

        # 테이블 전체 텍스트 (처음 3행)로 유형 판별
        sample_text = ""
        for row in table.rows[:3]:
            sample_text += " ".join(c.text.strip() for c in row.cells) + " "

        # ── 일정/담당자 테이블 ──
        if not data.review_date and ('투심' in sample_text or '일정' in sample_text):
            _extract_schedule_table(table, data)

        # ── 회사 개요 테이블 ──
        if not data.company_name and (
            '회사명' in sample_text or '대표이사' in sample_text
            or '대표자' in sample_text or '납입자본금' in sample_text
        ):
            _extract_company_table(table, data)

        # ── 투자조건 요약 테이블 ──
        if not data.fund_usage:
            _extract_investment_summary_table(table, data)

    # ── 주주현황 → 지분율 ──
    _extract_shareholder_table(tables, data)


def _extract_schedule_table(table, data: InvestmentReportData):
    """일정/담당자 테이블. 키워드 매칭."""
    _FIELD_MAP = {
        '투심 일자': 'review_date', '투심일자': 'review_date',
        '조합투심위': 'committee_date',
        '계약일': 'contract_date',
        '자금집행일': 'fund_date',
        '투자 계정': 'fund_name', '투자계정': 'fund_name',
        '발굴': 'discoverer',
        '심사': 'reviewer',
        '사후관리': 'post_manager',
    }
    for row in table.rows:
        cells = [cell.text.strip() for cell in row.cells]
        row_text = " ".join(cells)
        for keyword, attr in _FIELD_MAP.items():
            if keyword in row_text and not getattr(data, attr):
                # 값은 마지막 셀 (또는 키워드 다음 셀)
                val = cells[-1].strip()
                if val and keyword not in val:
                    setattr(data, attr, val)
                else:
                    for i, c in enumerate(cells):
                        if keyword in c and i + 1 < len(cells):
                            setattr(data, attr, cells[i + 1].strip())
                            break


def _extract_company_table(table, data: InvestmentReportData):
    """회사 개요 테이블. 다양한 필드명 지원."""
    _FIELD_MAP = {
        '회사명': 'company_name',
        '대표이사': 'representative', '대표자': 'representative',
        '납입자본금': 'paid_in_capital',
        '액면가': 'par_value',
        '설립일': 'establishment_date',
        '인력현황': 'num_employees', '종업원': 'num_employees', '종업원수': 'num_employees',
        '주요사업': 'business_description', '주요 사업': 'business_description',
    }

    for row in table.rows:
        cells = [cell.text.strip() for cell in row.cells]
        if len(cells) < 2:
            continue
        row_text = " ".join(cells)

        # 일반 필드 매칭
        for keyword, attr in _FIELD_MAP.items():
            if keyword in row_text and not getattr(data, attr):
                for i, c in enumerate(cells):
                    if keyword in c and i + 1 < len(cells):
                        val = cells[i + 1].replace('\xa0', '').strip()
                        if attr == 'representative':
                            val = val.replace(' ', '')
                            # 이름만 추출 (학력/경력이 붙어있을 경우)
                            m_name = re.match(r'([가-힣]{2,4})', val)
                            if m_name:
                                val = m_name.group(1)
                        setattr(data, attr, val)
                        break

        # 사업자등록번호 (패턴 매칭)
        if ('사업자' in row_text or '등록번호' in row_text) and not data.business_registration:
            for c in cells:
                m = re.search(r'(\d{3}-\d{2}-\d{5})', c)
                if m:
                    data.business_registration = m.group(1)
                    break

        # 주소 (다양한 표현)
        if not data.address and ('주소' in row_text or '본사' in row_text or '소재지' in row_text):
            for c in reversed(cells):
                if c and '주소' not in c and '본사' not in c and '소재지' not in c and len(c) > 5:
                    data.address = c.replace('\n', ' ').strip()
                    break


def _extract_investment_summary_table(table, data: InvestmentReportData):
    """투자조건 요약 테이블에서 사용용도, 위약벌 등 추출."""
    for row in table.rows:
        cells = [cell.text.strip() for cell in row.cells]
        if len(cells) < 2:
            continue
        label = cells[0]
        value = cells[1] if len(cells) > 1 else ""

        if '투자금 사용용도' in label and '위반' not in label and not data.fund_usage:
            data.fund_usage = value


def _extract_shareholder_table(tables, data: InvestmentReportData):
    """주주현황 테이블에서 케이런 지분율 추출."""
    if data.share_ratio:
        return

    for table in tables:
        total_post_shares = 0
        keiren_row_data = None

        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            row_text = " ".join(cells)

            # 합계행
            if '합계' in row_text:
                for c in cells:
                    m = re.match(r'^([\d,]+)$', c.strip())
                    if m:
                        val = int(m.group(1).replace(',', ''))
                        if val > total_post_shares:
                            total_post_shares = val

            # 케이런 행
            if '케이런' in row_text and ('주' in row_text or '%' in row_text or any(c.replace(',','').isdigit() for c in cells)):
                keiren_row_data = cells

        if keiren_row_data:
            # 주식수 추출
            shares_in_row = []
            for c in keiren_row_data:
                m = re.search(r'^([\d,]+)$', c.strip())
                if m:
                    val = int(m.group(1).replace(',', ''))
                    if val > 100:
                        shares_in_row.append(val)

            # 비율 추출
            ratios = []
            for c in keiren_row_data:
                m = re.search(r'(\d+\.?\d*)\s*%', c)
                if m:
                    ratios.append(m.group(1))

            # 직접 계산 우선
            if total_post_shares > 0 and shares_in_row:
                our_shares = max(shares_in_row)
                calculated = round(our_shares / total_post_shares * 100, 2)
                data.share_ratio = f"{calculated}%"
                return
            if ratios:
                data.share_ratio = ratios[-1] + "%"
                return


# ━━━━━━━━━━━━━━━ 본문 텍스트 추출 ━━━━━━━━━━━━━━━

def _extract_from_text(full_text: str, data: InvestmentReportData, doc=None):
    """본문에서 투자조건, 동반투자 등 추출."""

    # ── 투자금액 ──
    if not data.investment_amount:
        for pat in [
            r'총\s*([\d,]{7,})\s*원',
            r'([\d,]{7,})\s*원.*?규모',
            r'투자금액\s*[:：]?\s*([\d,]{7,})\s*원',
        ]:
            m = re.search(pat, full_text)
            if m:
                data.investment_amount = m.group(1) + "원"
                break

    # ── 투자단가 ──
    if not data.issue_price:
        for pat in [
            r'주당\s*([\d,]+)\s*원',
            r'투자단가\s*[:：]?\s*([\d,]+)\s*원',
        ]:
            m = re.search(pat, full_text)
            if m:
                data.issue_price = m.group(1) + "원"
                break

    # ── 주식수 ──
    if not data.total_shares:
        m = re.search(r'([\d,]+)\s*주를?\s*주당', full_text)
        if m:
            data.total_shares = m.group(1) + "주"

    # ── 투자방식 ──
    if not data.stock_type:
        m = re.search(r'(상환전환우선주|전환우선주|상환우선주|보통주)', full_text)
        if m:
            data.stock_type = m.group(1)

    # ── Pre/Post 기업가치 ──
    if not data.pre_value:
        m = re.search(r'Pre[- ]?[Vv]alue.*?(\d+)\s*억', full_text)
        if m:
            data.pre_value = m.group(1) + "억원"
    if not data.post_value:
        m = re.search(r'Post[- ]?[Vv]alue.*?(\d+)\s*억', full_text)
        if m:
            data.post_value = m.group(1) + "억원"

    # ── 존속기간/상환/Refixing ──
    if not data.duration:
        m = re.search(r'만기\s*(\d+)\s*년', full_text)
        if m:
            data.duration = m.group(1) + "년"
    if not data.redemption_terms:
        m = re.search(r'발행\s*(\d+)\s*년\s*후.*?상환.*?(?:YTM|연복리)\s*(\d+)\s*%', full_text)
        if m:
            data.redemption_terms = f"{m.group(1)}년후부터 상환청구 가능, 연복리 {m.group(2)}%"
    if not data.refixing_terms:
        m = re.search(r'IPO/M&A\s*(\d+)\s*%', full_text)
        if m:
            data.refixing_terms = f"IPO/M&A {m.group(1)}%"

    # ── 투자금 사용용도 (테이블 우선, 텍스트 fallback) ──
    if not data.fund_usage and doc:
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                if len(cells) >= 2 and '투자금 사용용도' in cells[0] and '위반' not in cells[0]:
                    data.fund_usage = cells[1].strip()
                    break
            if data.fund_usage:
                break
    if not data.fund_usage:
        for line in full_text.split('\n'):
            if any(kw in line for kw in ['운영자금', '설비투자', '연구개발', '시설자금']):
                if len(line) < 200:
                    data.fund_usage = line.strip()
                    break

    # ── 동반투자 ──
    if not data.co_investors:
        m = re.search(r'동반투자(?:기관|내역)[：:]?\s*(.+)', full_text)
        if m:
            pairs = re.findall(r'(\S+)\s+(\d+)억', m.group(1))
            for name, amount in pairs:
                data.co_investors.append((name, f"{amount}억원", ""))


# ━━━━━━━━━━━━━━━ PDF 전용 ━━━━━━━━━━━━━━━

def _extract_from_pdf_text(full_text: str, data: InvestmentReportData):
    """PDF 텍스트에서 테이블 없이 모든 데이터를 추출."""
    m = re.search(r'[㈜(주)]\s*(\S+)', full_text)
    if m:
        data.company_name = "㈜" + m.group(1)

    m = re.search(r'대표이사[:\s]*([가-힣]{2,4})', full_text)
    if m:
        data.representative = m.group(1)

    m = re.search(r'(\d{3}-\d{2}-\d{5})', full_text)
    if m:
        data.business_registration = m.group(1)

    m = re.search(r'(' + _REGIONS + r')[^\n]{5,80}', full_text)
    if m:
        data.address = m.group(0).strip()

    m = re.search(r'설립일[:\s]*([\d.]+)', full_text)
    if m:
        data.establishment_date = m.group(1)

    for pat in [
        r'((?:케이런|IBK)\S*\s*\S*\s*(?:투자)?조합)',
        r'(\S+\d+호\S*\s*(?:투자)?조합)',
        r'(20\d{2}\s*\S*\s*\S*\s*\d+호\s*펀드)',
    ]:
        m = re.search(pat, full_text)
        if m:
            data.fund_name = m.group(1)
            break

    m = re.search(r'발굴[:\s]*(\S+(?:\s*\(\d+%\))?)', full_text)
    if m:
        data.discoverer = m.group(1)
    m = re.search(r'심사[:\s]*(\S+(?:\s*\(\d+%\))(?:\s*,?\s*\S+(?:\s*\(\d+%\)))*)', full_text)
    if m:
        data.reviewer = m.group(1)
    m = re.search(r'사후관리[:\s]*(\S+(?:\s*\(\d+%\))?)', full_text)
    if m:
        data.post_manager = m.group(1)
