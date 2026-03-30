"""투자계약서에서 데이터를 추출하는 모듈. 키워드 기반 조문 자동 매칭."""
import re
from dataclasses import dataclass, field
from docx import Document


@dataclass
class InvestmentContractData:
    # 당사자 정보
    company_name: str = ""
    representative: str = ""
    address: str = ""
    interested_party: str = ""  # 이해관계인

    # 제1조 - 신주 발행 정보
    stock_type: str = ""           # 상환전환우선주 등
    total_shares: str = ""         # 발행 총수
    par_value: str = ""            # 액면가
    issue_price: str = ""          # 1주당 발행가액
    total_investment: str = ""     # 총 인수가액
    payment_date: str = ""         # 납입기일
    contract_date: str = ""        # 계약 체결일

    # 별지1 - 우선주 조건
    duration: str = ""             # 존속기간
    redemption_terms: str = ""     # 상환조건
    conversion_terms: str = ""     # 전환조건
    refixing_terms: str = ""       # Refixing 조건
    other_terms: str = ""          # 기타 (우선매수권, 공동매도참여권 등)

    # 별지3 - 투자금 사용용도
    fund_usage: str = ""

    # 조문 번호 (키워드 기반 자동 탐색)
    article_fund_usage: str = ""        # 투자금의 용도 및 제한
    article_consent: str = ""           # 동의권 및 협의권
    article_buyback: str = ""           # 주식매수청구권
    article_damages: str = ""           # 손해배상 및 위약벌
    article_delay_penalty: str = ""     # 지연배상금

    # 의무불이행 이자율
    buyback_rate: str = ""         # 주식매수청구권 이자율
    penalty_rate: str = ""         # 위약벌 비율
    delay_rate: str = ""           # 지연배상금 비율
    redemption_rate: str = ""      # 상환 이자율 (연복리)

    # 경고 메시지
    warnings: list = field(default_factory=list)


def extract_contract_data(filepath: str) -> InvestmentContractData:
    """투자계약서 DOCX에서 데이터를 추출한다."""
    doc = Document(filepath)
    paragraphs = [p.text.strip() for p in doc.paragraphs]
    full_text = "\n".join(paragraphs)
    data = InvestmentContractData()

    # --- 당사자 정보 추출 ---
    _extract_parties(paragraphs, data)

    # --- 제1조 신주 발행 정보 ---
    _extract_article1(full_text, data)

    # --- 계약 체결일 ---
    _extract_contract_date(full_text, data)

    # --- 키워드 기반 조문 번호 탐색 ---
    _extract_article_numbers(paragraphs, data)

    # --- 의무불이행 이자율 추출 ---
    _extract_penalty_rates(full_text, data)

    # --- 별지1 우선주 조건 ---
    _extract_appendix1(full_text, data)

    # --- 별지3 투자금 사용용도 ---
    _extract_appendix3(full_text, doc, data)

    return data


def _extract_parties(paragraphs: list, data: InvestmentContractData):
    """당사자 정보 (회사명, 대표이사, 주소) 추출."""
    in_company_section = False
    in_interested_section = False

    for i, p in enumerate(paragraphs):
        # "투자기업" 또는 "회사" 섹션 감지
        if re.match(r'2\.\s*투자기업', p) or p.startswith('2. 투자기업'):
            in_company_section = True
            in_interested_section = False
            continue
        if re.match(r'3\.\s*이해관계인', p):
            in_company_section = False
            in_interested_section = True
            continue
        if re.match(r'[4-9]\.\s', p) or p.startswith('제1장'):
            in_company_section = False
            in_interested_section = False

        if in_company_section:
            if '주식회사' in p and not data.company_name:
                data.company_name = p.strip()
            elif re.search(r'(서울|경기|인천|부산|대구|광주|대전|울산|세종|강원|충북|충남|전북|전남|경북|경남|제주)', p) and not data.address:
                data.address = p.strip()
            elif '대표이사' in p and not data.representative:
                name = re.sub(r'대표이사\s*', '', p).strip()
                data.representative = name

        if in_interested_section:
            if not data.interested_party and p and not re.match(r'\d+\.', p):
                name = p.strip()
                if len(name) <= 10 and name:
                    data.interested_party = name


def _extract_article1(full_text: str, data: InvestmentContractData):
    """제1조 신주 발행 정보 추출."""
    # 신주 종류
    m = re.search(r'본건 신주의 종류.*?:\s*별지', full_text)
    # 별지1에서 종류 추출
    m2 = re.search(r'1\.\s*본건 신주의 종류\s*\n\s*(.+)', full_text)
    if m2:
        raw = m2.group(1).strip()
        # "(이하 ...)" 부분 제거
        raw = re.sub(r'\(이하.*', '', raw).strip()
        # "상환전환우선주식" → "상환전환우선주" (끝의 "식" 제거)
        raw = re.sub(r'주식$', '주', raw)
        data.stock_type = raw
    if not data.stock_type:
        m3 = re.search(r'(상환전환우선주|전환우선주|상환우선주|보통주)', full_text[:500])
        if m3:
            data.stock_type = m3.group(1)

    # 발행 총수
    m = re.search(r'본건 신주의 발행 총수\s*:\s*([\d,]+)\s*주', full_text)
    if m:
        data.total_shares = m.group(1)

    # 액면가
    m = re.search(r'1주당 액면가액\s*:\s*금\s*([\d,]+)\s*원', full_text)
    if m:
        data.par_value = m.group(1)

    # 발행가액
    m = re.search(r'1주당 발행가액\s*:\s*금\s*([\d,]+)\s*원', full_text)
    if m:
        data.issue_price = m.group(1)

    # 총 인수가액
    m = re.search(r'총 인수가액\s*:\s*금\s*([\d,]+)\s*원', full_text)
    if m:
        data.total_investment = m.group(1)

    # 납입기일
    m = re.search(r'납입기일\s*:\s*(\d{4})년\s*(\d{1,2})월\s*\[?(\d{1,2})\]?\s*일', full_text)
    if m:
        data.payment_date = f"{m.group(1)}년 {m.group(2)}월 {m.group(3)}일"


def _extract_contract_date(full_text: str, data: InvestmentContractData):
    """계약 체결일 추출."""
    m = re.search(r'(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일.*?본 계약 체결일', full_text)
    if m:
        data.contract_date = f"{m.group(1)}년 {m.group(2)}월 {m.group(3)}일"
    else:
        m = re.search(r'본.*?계약.*?(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일', full_text[:500])
        if m:
            data.contract_date = f"{m.group(1)}년 {m.group(2)}월 {m.group(3)}일"


def _extract_article_numbers(paragraphs: list, data: InvestmentContractData):
    """키워드 기반으로 조문 번호를 자동 탐색한다."""
    keyword_map = {
        'article_fund_usage': ['투자금의 용도', '투자금의 용도 및 제한'],
        'article_consent': ['동의권 및 협의권', '동의권', '경영사항에 대한'],
        'article_buyback': ['주식매수청구권'],
        'article_damages': ['손해배상 및 위약벌', '손해배상', '위약벌'],
        'article_delay_penalty': ['지연배상금', '지연손해금'],
    }

    # 조문 제목 패턴: "제 N 조  제목" or "제N조 제목"
    article_pattern = re.compile(r'제\s*(\d+)\s*조\s+(.+)')

    for p in paragraphs:
        m = article_pattern.match(p)
        if not m:
            continue
        num = m.group(1)
        title = m.group(2).strip()

        for field_name, keywords in keyword_map.items():
            for kw in keywords:
                if kw in title:
                    current = getattr(data, field_name)
                    if not current:
                        setattr(data, field_name, num)
                    break


def _extract_penalty_rates(full_text: str, data: InvestmentContractData):
    """의무불이행 관련 이자율/비율 추출."""
    # 주식매수청구권 이자율: "연 12%의 이율" 패턴
    buyback_section = _find_section(full_text, '주식매수청구권')
    if buyback_section:
        m = re.search(r'연\s*(\d+)\s*%', buyback_section)
        if m:
            data.buyback_rate = m.group(1)

    # 위약벌: "투자금의 12%"
    damages_section = _find_section(full_text, '손해배상 및 위약벌')
    if damages_section:
        m = re.search(r'투자금의?\s*(\d+)\s*%', damages_section)
        if m:
            data.penalty_rate = m.group(1)

    # 지연배상금: "연 12%에 해당하는"
    delay_section = _find_section(full_text, '지연배상금')
    if delay_section:
        m = re.search(r'연\s*(\d+)\s*%', delay_section)
        if m:
            data.delay_rate = m.group(1)

    # 상환 이자율: "연 복리 5%"
    m = re.search(r'상환.*?연\s*복리\s*(\d+)\s*%', full_text)
    if m:
        data.redemption_rate = m.group(1)


def _extract_appendix1(full_text: str, data: InvestmentContractData):
    """별지1 우선주 조건 추출."""
    # 존속기간
    m = re.search(r'존속기간.*?(\d+)\s*년', full_text)
    if m:
        data.duration = f"{m.group(1)}년"

    # 상환조건
    m = re.search(r'효력발생일로부터\s*(\d+)\s*년이?\s*경과한\s*날로부터.*?상환', full_text)
    if m:
        years = m.group(1)
        rate = data.redemption_rate or ""
        data.redemption_terms = f"투자 {years}년후부터 연복리 {rate}%로 상환청구가능" if rate else f"투자 {years}년후부터 상환청구가능"

    # 전환조건
    conversion_parts = []
    if re.search(r'효력발생일부터.*?존속기간 만료일까지.*?전환', full_text):
        conversion_parts.append("발행일 익일부터 만기까지")
    if re.search(r'존속기간.*?만료일.*?보통주식으로 전환', full_text):
        conversion_parts.append("존속기간 이후 보통주 자동 전환")
    m = re.search(r'종류주식\s*1주.*?보통주식.*?(\d+)\s*주', full_text)
    if m:
        conversion_parts.append(f"우선주 1주당 보통주 {m.group(1)}주로 전환")
    data.conversion_terms = ", ".join(conversion_parts) if conversion_parts else ""

    # Refixing
    refixing_parts = []
    m = re.search(r'공모단가.*?(\d+)\s*%', full_text)
    if m:
        refixing_parts.append(f"IPO/M&A {m.group(1)}%")
    if re.search(r'전환가격보다 낮은 발행가격', full_text):
        refixing_parts.append("투자단가 이하 유상증자 등")
    data.refixing_terms = "Refixing: " + ", ".join(refixing_parts) if refixing_parts else ""

    # 기타 (우선매수권, 공동매도참여권)
    other_parts = []
    if re.search(r'우선매수권', full_text):
        other_parts.append("우선매수권")
    if re.search(r'공동매도참여권|공동매도권', full_text):
        other_parts.append("공동매도참여권")
    data.other_terms = ", ".join(other_parts) if other_parts else ""


def _extract_appendix3(full_text: str, doc: Document, data: InvestmentContractData):
    """별지3 투자금 사용용도 추출."""
    # 테이블에서 투자금 사용용도 찾기
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            joined = " ".join(cells)
            if '사용용도' in joined or '사용목적' in joined:
                for cell_text in cells:
                    if cell_text and '사용용도' not in cell_text and '사용목적' not in cell_text:
                        if len(cell_text) > 5:
                            data.fund_usage = cell_text
                            return

    # 텍스트에서 추출 시도
    m = re.search(r'별지\s*3.*?투자금.*?사용.*?용도\s*\n(.+)', full_text)
    if m:
        data.fund_usage = m.group(1).strip()


def _find_section(full_text: str, keyword: str) -> str:
    """키워드를 포함하는 조문 섹션의 텍스트를 반환한다."""
    pattern = re.compile(r'제\s*\d+\s*조\s+.*?' + re.escape(keyword) + r'.*?\n', re.DOTALL)
    m = pattern.search(full_text)
    if m:
        start = m.start()
        # 다음 조문까지의 텍스트
        next_article = re.search(r'\n제\s*\d+\s*조\s+', full_text[m.end():])
        if next_article:
            return full_text[start:m.end() + next_article.start()]
        return full_text[start:start + 3000]
    return ""
