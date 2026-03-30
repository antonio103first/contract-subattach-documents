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
    """투자계약서에서 데이터를 추출한다. DOCX와 PDF 모두 지원."""
    ext = filepath.lower().rsplit('.', 1)[-1] if '.' in filepath else ''

    if ext == 'pdf':
        from extractors.pdf_extractor import extract_text_from_pdf
        full_text = extract_text_from_pdf(filepath)
        paragraphs = [p.strip() for p in full_text.split('\n') if p.strip()]
        doc = None
    else:
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
    full = "\n".join(paragraphs)

    # 패턴 1: "2. 투자기업" 섹션 (위밋 스타일)
    # 패턴 2: "발행회사 : 주식회사 XXX" (에이치투 스타일)
    # 패턴 3: "2. 피투자자 :" 섹션

    # 회사명 추출
    for pattern in [
        r'발행회사\s*[:：]\s*(주식회사\s*\S+)',
        r'피투자자\s*[:：]?\s*\n?\s*(주식회사\s*\S+)',
        r'2\.\s*투자기업.*?\n.*?(주식회사\s*\S+)',
        r'"회사"의 상호\s*[:：]\s*(주식회사\s*\S+)',
    ]:
        m = re.search(pattern, full)
        if m and not data.company_name:
            data.company_name = m.group(1).strip().rstrip('"').rstrip("'")
            break

    # 대표이사 추출 - 피투자자/회사 섹션에서 추출 (투자자 대표가 아님)
    # 패턴: "2. 피투자자:" 섹션 내의 대표이사
    company_section = ""
    m_company = re.search(r'(2\.\s*피투자자.*?)(?=3\.\s*이해관계인)', full, re.DOTALL)
    if m_company:
        company_section = m_company.group(1)
    else:
        m_company = re.search(r'(2\.\s*투자기업.*?)(?=3\.\s*이해관계인)', full, re.DOTALL)
        if m_company:
            company_section = m_company.group(1)

    if company_section:
        m = re.search(r'대표이사\s*[:：]?\s*([가-힣]{2,4})', company_section)
        if m:
            data.representative = m.group(1).strip()

    # fallback: 이해관계인 섹션에서 대표이사
    if not data.representative:
        m = re.search(r'이해관계인.*?대표이사\s*([가-힣]{2,4})', full)
        if m:
            data.representative = m.group(1).strip()

    # 주소 추출 (피투자자/회사 섹션)
    in_company_section = False
    for i, p in enumerate(paragraphs):
        if re.match(r'2\.\s*(투자기업|피투자자)', p) or '발행회사' in p:
            in_company_section = True
            continue
        if re.match(r'3\.\s*(이해관계인|투자)', p) or '이해관계인' in p:
            if in_company_section and not data.address:
                # 이 섹션 끝
                pass
            in_company_section = False

        if in_company_section:
            if re.search(r'주소\s*[:：]', p):
                addr = re.sub(r'주소\s*[:：]\s*', '', p).strip()
                if addr:
                    data.address = addr
            elif not data.address and re.search(r'(서울|경기|인천|부산|대구|광주|대전|울산|세종|강원|충북|충남|전북|전남|경북|경남|제주)', p):
                data.address = p.strip()

    # 이해관계인
    m = re.search(r'이해관계인.*?[:：]\s*.*?대표이사\s*([가-힣]{2,4})', full)
    if m:
        data.interested_party = m.group(1)


def _extract_article1(full_text: str, data: InvestmentContractData):
    """신주 발행 정보 추출 (제1조 또는 제5조 등 구조에 무관하게)."""
    # 신주 종류 - 여러 패턴 시도
    for pattern in [
        r'1\.\s*본건 신주의 종류\s*\n\s*(.+)',
        r'종류와 수\s*:\s*기명식\s*(\S+우선주)\s',
        r'(상환전환우선주식?|전환우선주식?|상환우선주식?)',
    ]:
        m = re.search(pattern, full_text)
        if m and not data.stock_type:
            raw = m.group(1).strip()
            raw = re.sub(r'\(이하.*', '', raw).strip()
            raw = re.sub(r'주식$', '주', raw)
            data.stock_type = raw
            break
    if not data.stock_type:
        m = re.search(r'(상환전환우선주|전환우선주|상환우선주|보통주)', full_text[:1000])
        if m:
            data.stock_type = m.group(1)

    # 발행 총수 - 여러 패턴
    for pattern in [
        r'본건 신주의 발행 총수\s*:\s*([\d,]+)\s*주',
        r'종류와 수\s*:.*?([\d,]+)\s*주',
        r'"본 주식"의 종류와 수\s*:.*?([\d,]+)\s*주',
    ]:
        m = re.search(pattern, full_text)
        if m and not data.total_shares:
            data.total_shares = m.group(1)
            break

    # 액면가
    for pattern in [
        r'1주당 액면가액\s*:\s*금\s*([\d,]+)\s*원',
        r'1주의 금액.*?:\s*금\s*([\d,]+)\s*원',
        r'액면가[)]\s*:\s*금\s*([\d,]+)\s*원',
    ]:
        m = re.search(pattern, full_text)
        if m and not data.par_value:
            data.par_value = m.group(1)
            break

    # 발행가액 (1주당)
    for pattern in [
        r'1주당 발행가액\s*:\s*금\s*([\d,]+)\s*원',
        r'1주당 발행가액.*?:\s*금\s*([\d,]+)\s*원',
        r'발행가액.*?인수가액.*?:\s*금\s*([\d,]+)\s*원',
    ]:
        m = re.search(pattern, full_text)
        if m and not data.issue_price:
            data.issue_price = m.group(1)
            break

    # 총 인수가액/인수대금
    for pattern in [
        r'총 인수가액\s*:\s*금\s*([\d,]+)\s*원',
        r'총 인수대금\s*:.*?\\?([\d,]+)\s*원\)',
        r'총 인수대금\s*:.*?([\d,]+)\s*원',
        r'\\([\d,]{7,})\s*원\)',
    ]:
        m = re.search(pattern, full_text)
        if m and not data.total_investment:
            data.total_investment = m.group(1)
            break

    # 납입기일
    m = re.search(r'납입기일\s*:\s*(\d{4})년\s*(\d{1,2})월\s*\[?\s*(\d{1,2})?\s*\]?\s*일', full_text)
    if m:
        day = m.group(3) or ""
        data.payment_date = f"{m.group(1)}년 {m.group(2)}월 {day}일".strip()


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
