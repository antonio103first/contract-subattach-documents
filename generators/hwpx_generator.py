"""준법사항체크리스트 HWPX 파일 생성 모듈.
실제 양식 HWPX를 복사하여 section0.xml의 placeholder를 치환.
작성지침에 따라 표1~5를 자동으로 채운다."""
import os
import re
import zipfile
from datetime import datetime, date


DEFAULT_TEMPLATE = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "(양식) 준법사항체크리스트(안)_2024 IBK혁신 케이런 모빌리티 7호 펀드_251208.hwpx"
)


def generate_hwpx_checklist(contract_data, report_data, output_path: str,
                             template_path: str = None):
    """양식 HWPX를 복사하여 준법사항체크리스트를 생성한다."""
    template_path = template_path or DEFAULT_TEMPLATE
    if not os.path.exists(template_path):
        print(f"[ERROR] HWPX 양식 파일을 찾을 수 없습니다: {template_path}")
        return

    cd = contract_data
    rd = report_data
    replacements = _build_all_replacements(cd, rd)
    warnings = _check_mismatches(cd, rd)

    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    _copy_and_replace(template_path, output_path, replacements)

    print(f"[OK] HWPX 준법사항체크리스트 생성 완료: {output_path}")
    if warnings:
        print(f"\n[주의] HWPX {len(warnings)}건의 불일치 발견:")
        for note in warnings:
            print(f"  - {note}")


def _build_all_replacements(cd, rd) -> dict:
    """작성지침에 따라 모든 placeholder → 실제 값 매핑을 구성."""

    # ── 기본 데이터 준비 ──
    company = rd.company_name or cd.company_name
    short = company
    if '주식회사' in company:
        short = "㈜" + company.replace('주식회사', '').replace('㈜', '').strip()
    if not short.startswith('㈜') and '㈜' not in short and '(주)' not in short:
        short = "㈜" + short

    rep = rd.representative or cd.representative or ""
    addr = rd.address or cd.address or ""
    biz_id = rd.business_registration or "(미확인)"
    stock_type = cd.stock_type or rd.stock_type or ""
    inv_amt = _fmt_won(cd.total_investment)
    iss_price = _fmt_won(cd.issue_price)
    ratio = rd.share_ratio or ""
    discoverer = rd.discoverer or ""
    reviewer = rd.reviewer or ""
    post_mgr = rd.post_manager or ""

    # ── 표2: 주요 투자조건 ──
    cond_parts = []
    if cd.duration:
        cond_parts.append(f" - 존속기간 : {cd.duration}")
    if cd.redemption_terms:
        cond_parts.append(f" - 상환조건 : {cd.redemption_terms}")
    if cd.conversion_terms:
        cond_parts.append(f" - 전환조건 : {cd.conversion_terms}")
    extras = []
    if cd.refixing_terms:
        extras.append(cd.refixing_terms.replace("Refixing: ", ""))
    if cd.other_terms:
        extras.append(cd.other_terms)
    if extras:
        cond_parts.append(f" - 기타 : {', '.join(extras)} 등")
    cond_text = " ".join(cond_parts)

    # ── 표2: 위약벌 ──
    pen_parts = []
    if cd.penalty_rate:
        pen_parts.append(f" - 위약벌 : 투자금의 {cd.penalty_rate}%")
    if cd.delay_rate:
        pen_parts.append(f" - 지연배상금 : 실제 지급일까지 연 {cd.delay_rate}%")
    if cd.buyback_rate:
        pen_parts.append(f" - 주식매수청구권 : 투자원금 및 {cd.buyback_rate}%")
    pen_text = " ".join(pen_parts)

    # ── 표4: 벤처기업 해당여부 (적/부 판단) ──
    # 창업기업: 설립일 기준 7년 이내
    startup_yn = _check_startup(rd.establishment_date)
    estab_str = rd.establishment_date or "0000년 00월 00일"
    # 벤처기업
    venture_yn = "적(Y)" if rd.is_venture == "Y" else "부(N)"
    # 이노비즈
    innobiz_yn = "적(Y)" if rd.is_innobiz == "Y" else "부(N)"

    # ── 표5: 준법사항 (적/부 판단) ──
    # 이해관계인 (계약서에서)
    interested = cd.interested_party or cd.representative or rep
    # 산업분류코드
    ind_code = rd.industry_code or "(확인 필요)"
    ind_desc = rd.business_description or ""
    # 주목적투자 해당여부
    purpose_transport = _yn(rd.purpose_transport)
    purpose_mobility = _yn(rd.purpose_mobility)
    purpose_south = _yn(rd.purpose_south)
    purpose_tcb = _yn(rd.purpose_tcb)
    # 투자구분 (신규/구주)
    is_new_stock = "신규" in (rd.investment_type or "") or "신주" in (cd.stock_type or "")
    # 해외투자 여부 (국내 주소면 부)
    is_domestic = bool(re.search(r'서울|경기|인천|부산|대구|광주|대전|울산|세종|강원|충|전|경|제주', addr))
    # 투자기간 이내 (2029.9.8 이전)
    invest_in_period = "적" if _is_before_deadline() else "부"

    return {
        # ── 단순 치환 (유일 placeholder) ──
        '_simple': {
            '㈜AAA': short,
            '000-00-00000': biz_id,
            # 표4 설립일
            '0000년 00월 00일': _format_estab_date(estab_str),
            # 표5 산업분류코드
            '한국표준산업분류코드 :': f'한국표준산업분류코드 : ({ind_code}) {ind_desc}',
            # 표5 이해관계인
            '이해관계인 :': f'이해관계인 : {interested}',
            # 표5 투자기간 종료일 (첫 번째 빈칸)
            '년  월  일': '2029년  9월  8일',
        },
        # ── 표2 주요 투자조건 (개별 <hp:t> 태그 치환) ──
        '_conditions': {
            ' - 존속기간 :': f' - 존속기간 : {cd.duration}' if cd.duration else ' - 존속기간 :',
            ' - 상환조건 :': f' - 상환조건 : {cd.redemption_terms}' if cd.redemption_terms else ' - 상환조건 :',
            ' - 전환조건 :': f' - 전환조건 : {cd.conversion_terms}' if cd.conversion_terms else ' - 전환조건 :',
            ' - 기타 :': f' - 기타 : {", ".join(extras)} 등' if extras else ' - 기타 :',
            ' - 위약벌 :': f' - 위약벌 : 투자금의 {cd.penalty_rate}%' if cd.penalty_rate else ' - 위약벌 :',
            ' - 지연배상금 :': f' - 지연배상금 : 실제 지급일까지 연 {cd.delay_rate}%' if cd.delay_rate else ' - 지연배상금 :',
            ' - 주식매수청구권 :': f' - 주식매수청구권 : 투자원금 및 {cd.buyback_rate}%' if cd.buyback_rate else ' - 주식매수청구권 :',
        },
        # ── 순서 기반 치환 (OOO 4회: 대표→발굴→심사→사후관리) ──
        '_ordered': {
            'OOO': [rep, discoverer, reviewer, post_mgr],
            'OO': [addr],
        },
        # ── 표4 적/부 치환 (순서: 창업→벤처→이노비즈) ──
        '_yn_markers': [
            startup_yn,   # 창업기업
            venture_yn,   # 벤처기업
            innobiz_yn,   # 이노비즈
        ],
        # ── 표5 데이터 (별도 처리) ──
        '_table5': {
            'interested': interested,
            'industry_code': ind_code,
            'industry_desc': ind_desc,
            'is_new_stock': is_new_stock,
            'is_domestic': is_domestic,
            'invest_in_period': invest_in_period,
            'purpose_transport': purpose_transport,
            'purpose_mobility': purpose_mobility,
            'purpose_south': purpose_south,
            'purpose_tcb': purpose_tcb,
            'purpose_tcb_detail': rd.purpose_tcb_detail or "",
        },
    }


# ━━━━━━━━━━━━━━━ 치환 엔진 ━━━━━━━━━━━━━━━

def _apply_replacements(text: str, replacements: dict) -> str:
    """모든 치환 규칙을 적용."""
    simple = replacements.get('_simple', {})
    ordered = replacements.get('_ordered', {})
    conditions = replacements.get('_conditions', {})
    yn_markers = replacements.get('_yn_markers', [])

    # 1. 단순 치환
    for old_val, new_val in simple.items():
        if old_val and new_val:
            text = text.replace(old_val, _xml_safe(new_val))

    # 1.5. 조건/위약벌 개별 태그 치환
    for old_val, new_val in conditions.items():
        if old_val and new_val:
            text = text.replace(old_val, _xml_safe(new_val))

    # 2. 순서 기반 치환 (XML과 PrvText 모두)
    for placeholder, values in ordered.items():
        for new_val in values:
            if new_val:
                safe_val = _xml_safe(new_val)
                # XML: >OOO<
                pat = '>' + re.escape(placeholder) + '<'
                repl = '>' + safe_val + '<'
                text = re.sub(pat, repl, text, count=1)
                # PrvText: <OOO>
                pat2 = '<' + re.escape(placeholder) + '>'
                repl2 = '<' + safe_val + '>'
                text = re.sub(pat2, repl2, text, count=1)

    # 3. 표4 적(Y)/부(N) 치환 - 순서대로 처리
    for yn_val in yn_markers:
        if '적' in yn_val:
            # "적(Y)  부(N)" → "적(Y)" (부 삭제)
            text = text.replace('적(Y)  부(N)', '적(Y)', 1)
        else:
            # "적(Y)  부(N)" → "부(N)" (적 삭제)
            text = text.replace('적(Y)  부(N)', '부(N)', 1)

    return text


def _copy_and_replace(template_path: str, output_path: str, replacements: dict):
    """양식 HWPX를 복사하고 치환. flag_bits 등 원본 메타데이터 완전 보존."""
    import struct
    import shutil

    temp_path = output_path + '.tmp'

    with zipfile.ZipFile(template_path, 'r') as zin:
        modified = {}
        for fname in ('Contents/section0.xml', 'Preview/PrvText.txt'):
            text = zin.read(fname).decode('utf-8', errors='replace')
            text = _apply_replacements(text, replacements)
            modified[fname] = text.encode('utf-8')

        with zipfile.ZipFile(temp_path, 'w') as zout:
            for item in zin.infolist():
                if item.filename in modified:
                    data = modified[item.filename]
                else:
                    data = zin.read(item.filename)
                zout.writestr(item, data)

    # flag_bits 복원: ZIP 바이너리에서 직접 패치
    _patch_flag_bits(template_path, temp_path)
    os.replace(temp_path, output_path)


def _patch_flag_bits(template_path: str, target_path: str):
    """원본 HWPX의 flag_bits를 생성본에 복사. (local + central directory 모두)"""
    import struct

    with zipfile.ZipFile(template_path) as zt:
        orig_flags = {item.filename: item.flag_bits for item in zt.infolist()}

    # ZIP 파일 바이너리 읽기
    with open(target_path, 'rb') as f:
        data = bytearray(f.read())

    # Local file headers: signature = PK\x03\x04
    offset = 0
    while offset < len(data) - 4:
        sig = struct.unpack_from('<I', data, offset)[0]
        if sig == 0x04034b50:  # Local file header
            fname_len = struct.unpack_from('<H', data, offset + 26)[0]
            extra_len = struct.unpack_from('<H', data, offset + 28)[0]
            fname = data[offset + 30: offset + 30 + fname_len].decode('utf-8', errors='replace')
            if fname in orig_flags:
                struct.pack_into('<H', data, offset + 6, orig_flags[fname])
            comp_size = struct.unpack_from('<I', data, offset + 18)[0]
            offset += 30 + fname_len + extra_len + comp_size
        elif sig == 0x02014b50:  # Central directory header
            fname_len = struct.unpack_from('<H', data, offset + 28)[0]
            extra_len = struct.unpack_from('<H', data, offset + 30)[0]
            comment_len = struct.unpack_from('<H', data, offset + 32)[0]
            fname = data[offset + 46: offset + 46 + fname_len].decode('utf-8', errors='replace')
            if fname in orig_flags:
                struct.pack_into('<H', data, offset + 8, orig_flags[fname])
            offset += 46 + fname_len + extra_len + comment_len
        elif sig == 0x06054b50:  # End of central directory
            break
        else:
            offset += 1

    with open(target_path, 'wb') as f:
        f.write(data)


# ━━━━━━━━━━━━━━━ 유틸 ━━━━━━━━━━━━━━━

def _xml_safe(text: str) -> str:
    """XML 특수문자를 이스케이프한다. 이미 이스케이프된 것은 건너뜀."""
    if not text:
        return text
    # 이미 이스케이프된 &amp; 등은 보존
    text = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9a-fA-F]+;)', '&amp;', text)
    # < > 는 XML 태그가 아닌 경우만 (PrvText에서는 구분자로 사용)
    # section0.xml에서는 이미 태그 안이므로 < > 치환 불필요
    return text


def _fmt_won(val: str) -> str:
    if not val:
        return ""
    v = val.replace(',', '').replace('원', '').strip()
    if v:
        return f"{int(v):,}원"
    return val


def _check_startup(establishment_date: str) -> str:
    """설립일 기준 7년 이내인지 확인."""
    if not establishment_date:
        return "부(N)"
    try:
        # "2017.05.25" or "2017년 5월 25일" 등
        cleaned = re.sub(r'[년월일\s]', '.', establishment_date).strip('.')
        parts = [p for p in cleaned.split('.') if p]
        if len(parts) >= 1:
            year = int(parts[0])
            current_year = datetime.now().year
            if current_year - year <= 7:
                return "적(Y)"
    except (ValueError, IndexError):
        pass
    return "부(N)"


def _format_estab_date(estab: str) -> str:
    """설립일을 "YYYY년 MM월 DD일" 형태로 변환."""
    if not estab or estab == "0000년 00월 00일":
        return "0000년 00월 00일"
    cleaned = re.sub(r'[년월일\s]', '.', estab).strip('.')
    parts = [p for p in cleaned.split('.') if p]
    if len(parts) >= 3:
        return f"{parts[0]}년 {parts[1]}월 {parts[2]}일"
    elif len(parts) == 2:
        return f"{parts[0]}년 {parts[1]}월"
    return estab


def _yn(val: str) -> str:
    """'해당'/'미해당' 등을 '적'/'부'로 변환."""
    if not val:
        return "부"
    v = val.strip()
    if v in ('해당', '가능', '적합', 'Y', 'O', '있음'):
        return "적"
    return "부"


def _is_before_deadline() -> bool:
    """현재 날짜가 투자기간 종료일(2029.9.8) 이전인지."""
    return date.today() < date(2029, 9, 8)


def _check_mismatches(cd, rd) -> list:
    warnings = []
    def _norm(v):
        return re.sub(r'[\s,원주%㈜주식회사(주)]', '', str(v or ''))
    checks = [
        ("투자금액", rd.investment_amount, _fmt_won(cd.total_investment)),
        ("투자단가", rd.issue_price, _fmt_won(cd.issue_price)),
        ("투자방식", rd.stock_type, cd.stock_type),
    ]
    for name, rv, cv in checks:
        if rv and cv and _norm(rv) != _norm(cv):
            warnings.append(f"{name}: 투심보고서={rv}, 투자계약서={cv}")
            print(f"[WARNING] {name} 불일치: 투심보고서={rv}, 투자계약서={cv}")
    return warnings
