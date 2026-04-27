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
    from extractors.contract_extractor import normalize_company_name
    company = rd.company_name or cd.company_name
    short = normalize_company_name(company) or company

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

    # ── 표5 준법사항: 적/부 순서 목록 (양식의 실제 칼럼 빈 셀 순서) ──
    committee_date = rd.committee_date or "(확인 필요)"

    # 투자방법 판단
    stock_lower = (cd.stock_type or rd.stock_type or "").replace(" ", "")
    is_stock_type = any(kw in stock_lower for kw in ['보통주', '우선주', 'RCPS'])
    is_cb_bw = any(kw in stock_lower for kw in ['CB', 'BW', '전환사채', '신주인수권'])

    table5_yn = [
        # ── 법령상 투자제한 (12개) ──
        "부",    # 자기 또는 제3자의 이익을 위한 조합 재산 사용 여부
        "부",    # 투자기업의 상호출자제한기업집단 소속 여부
        "부",    # 투자 제한업종 해당 여부
        "부",    # 취득 대상이 금융회사 등 주식 또는 지분인지 여부
        "부",    # 취득 대상이 이해관계인이 발행하거나 소유한 주식
        "부",    # 이해관계인에 대한 신용공여 행위 여부
        "부",    # 조합 명의로 제3자를 위하여 주식 취득/자금 중개
        "부",    # 조합이 투자한 업체로부터 차입
        "부",    # 투자계약서에 기재된 조건 외에 별도 투자조건 설정
        "적" if is_domestic else "부",  # 해외투자 요건 준수 여부
        "적",    # 2개 이상 기업 프로젝트 (확인 필요)
        "부",    # 기타 법령 위반 여부
        # ── 규약상 투자제한 ──
        "적",    # 제34조 제1항의 법상 의무투자 해당여부
        purpose_transport,   # 제35조 제1항 제1호 (국토교통분야)
        purpose_mobility,    # 제35조 제1항 제2호 (혁신성장 모빌리티)
        purpose_south,       # 제35조 제1항 제3호 (남부권 전략산업)
        "적",                # 제61조 제1항 제1호 (IBK 기업거래)
        purpose_tcb,         # 제61조 제1항 제2호 (TCB Ti-6 등급)
        "적",    # 제34조 제3항 동일기업 동일 프로젝트
        "적",    # 제34조 제4항 후행투자 (담당자 확인)
        "부" if is_new_stock else "적",  # 제34조 제2항 구주
        "부" if is_domestic else "적",   # 제34조 제2항 해외투자
        "부",    # 제34조 제8항 자금 대여 방식
        "부",    # 제34조 제10항 금지행위 (담당자 확인)
        invest_in_period,  # 제4조 제26호 투자기간 이내
        "적",    # 제8조 제5항 납입금액 충족
        "부",    # 제34조의2 이해상충 검토
        "적",    # 제37조 투자심의위원회 부의
        "적",    # 제61조 제14항 볼커룰
    ]

    # ── 투자방법 (O) 체크 ──
    # 양식의 (   ) 5개: 신규주식, 무담보사채, 조건부, 창업자주식, 프로젝트
    invest_method_checks = [
        is_stock_type,   # 신규로 발행되는 주식의 인수
        is_cb_bw,        # 무담보전환사채 등
        False,           # 조건부지분인수계약
        False,           # 개인/개인투자조합 3년 이상 보유 창업자 주식
        False,           # 프로젝트 투자
    ]

    # ── 비고란 적색 주석 (colAddr=3 빈 셀에 순서대로 삽입) ──
    legal_bigo = [
        "",                          # row2: 자기/제3자
        "[확인 필요: 중소기업 여부]",  # row3: 상호출자
        "",                          # row8: 조합명의
        "",                          # row9: 차입
        "",                          # row10: 별도조건
        "[별도 확인 필요]",            # row12: 프로젝트
    ]
    regulatory_bigo = [
        "",                          # row11: 해외투자
        "",                          # row12
        "",                          # row13
        "",                          # row14: 금지행위 (전체 적색 처리)
        "",                          # row15: 투자기간
        "",                          # row16: 납입금액
        "",                          # row17: 이해상충 (전체 적색 처리)
        "",                          # row18: 투심위
    ]

    # ── 담당자 확인 필요 → 해당 행 전체를 적색으로 표시할 항목 키워드 ──
    # 이 항목들은 행 전체(내용+평가+비고) 적색 표시
    red_full_rows = [
        '제34조 제4항의 후행투자 여부',
        '제34조 제10항에 의한 금지행위 여부',
        '제34조의 2 제1항의 이해상충여부 검토 여부',
        '제61조 제14항의 볼커룰',
        '제61조 제1항 제1호의 투자 해당여부',  # 투자의무4
    ]

    # ── 투자의무4 비고란 주석 ──
    # 제61조 제1항 제1호 → "별도 확인 필요" 적색 표시
    # (이 항목은 비고란이 이미 텍스트가 있으므로 텍스트 기반 주석으로 처리)

    # ── 텍스트 기반 주석 ──
    red_notes = {
        '(상세하게 발굴경위 기재)': rd.discovery_background or '(확인 필요)',
    }
    # 10번: TCB 등급 비고란 - "본건 TCB 등급: TI-" 뒤에 실제 등급 삽입
    tcb_detail = rd.purpose_tcb_detail or ""
    if tcb_detail:
        # "TI-3 등급(2025.8.28 발급)" → "본건 TCB 등급: TI-3 등급(2025.8.28 발급)"
        # 양식에서 "본건 TCB　등급: TI-" 다음의 빈 부분과 "0000.00.00 발급" 치환
        m = re.search(r'(TI-\d+)\s*등급.*?\(([\d.]+)\s*발급\)', tcb_detail)
        if m:
            red_notes['본건 TCB\u3000등급: TI-'] = f'본건 TCB 등급: {m.group(1)} 등급'
            red_notes['0000.00.00 발급'] = f'{m.group(2)} 발급'
        else:
            red_notes['본건 TCB\u3000등급: TI-'] = f'본건 TCB 등급: {tcb_detail}'
    # 투자의무4는 전체 적색 처리 (별도 확인 필요 삭제)
    # 투심위 예정일
    red_notes['년  월 일'] = committee_date

    return {
        '_simple': {
            '㈜AAA': short,
            '000-00-00000': biz_id,
            '0000년 00월 00일': _format_estab_date(estab_str),
            '한국표준산업분류코드 :': f'한국표준산업분류코드 : ({ind_code}) {ind_desc}',
            '이해관계인 :': f'이해관계인 : {interested}',
            '년  월  일': '2029년  9월  8일',
        },
        '_conditions': {
            # 표2 투자유형/금액/단가/지분율 (순서 기반 - >원< 치환)
            ' - 존속기간 :': f' - 존속기간 : {cd.duration}' if cd.duration else ' - 존속기간 :',
            ' - 상환조건 :': f' - 상환조건 : {cd.redemption_terms}' if cd.redemption_terms else ' - 상환조건 :',
            ' - 전환조건 :': f' - 전환조건 : {cd.conversion_terms}' if cd.conversion_terms else ' - 전환조건 :',
            ' - 기타 :': f' - 기타 : {", ".join(extras)} 등' if extras else ' - 기타 :',
            ' - 위약벌 :': f' - 위약벌 : 투자금의 {cd.penalty_rate}%' if cd.penalty_rate else ' - 위약벌 :',
            ' - 지연배상금 :': f' - 지연배상금 : 실제 지급일까지 연 {cd.delay_rate}%' if cd.delay_rate else ' - 지연배상금 :',
            ' - 주식매수청구권 :': f' - 주식매수청구권 : 투자원금 및 {cd.buyback_rate}%' if cd.buyback_rate else ' - 주식매수청구권 :',
        },
        '_ordered': {
            'OOO': [rep, discoverer, reviewer, post_mgr],
            'OO': [addr],
        },
        # 표2 투자유형/금액/단가/지분율: >원< 과 >%< 순서 치환 (4개씩)
        '_table2_values': {
            # text node 22~24: 첫 번째 행 (투자유형 행)
            # text node 27~29: 합계 행
            'stock_type': stock_type,
            'inv_amt': inv_amt,
            'iss_price': iss_price,
            'ratio': ratio,
        },
        '_yn_markers': [startup_yn, venture_yn, innobiz_yn],
        '_table5_yn': table5_yn,
        '_invest_method_checks': invest_method_checks,
        '_bigo_notes': legal_bigo + regulatory_bigo,
        '_red_full_rows': red_full_rows,  # 전체 적색 표시할 행 키워드
        '_red_notes': red_notes,
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

    # 3. 표2 투자유형/금액/단가/지분율 치환
    t2 = replacements.get('_table2_values', {})
    if t2.get('stock_type'):
        # 구분 컬럼 첫 번째 빈 셀 (text node 16 다음의 빈 셀)에 투자유형 삽입
        # 양식에서 구분 행 >원< 앞에 빈 셀이 있음. >< 패턴으로 치환은 어려우므로
        # >원< 자체를 값+원 으로 치환 (순서: 투자금액, 투자단가, 합계금액, 합계단가)
        pass

    if t2.get('inv_amt'):
        # 첫 번째 >원< → 투자금액
        text = re.sub(r'>원<', '>' + _xml_safe(t2['inv_amt']) + '<', text, count=1)
    if t2.get('iss_price'):
        # 두 번째 >원< → 투자단가
        text = re.sub(r'>원<', '>' + _xml_safe(t2['iss_price']) + '<', text, count=1)
    if t2.get('ratio'):
        # 첫 번째 >%< → 지분율
        text = re.sub(r'>%<', '>' + _xml_safe(t2['ratio']) + '<', text, count=1)
    if t2.get('stock_type'):
        # 기타(    ) 앞의 빈 셀에 투자유형 삽입은 어려우므로
        # "기타(    )" 를 투자유형으로 치환
        text = text.replace('기타(    )', _xml_safe(t2['stock_type']), 1)

    # 합계 행: 나머지 >원< 과 >%< 치환 (합계 = 같은 값)
    if t2.get('inv_amt'):
        text = re.sub(r'>원<', '>' + _xml_safe(t2['inv_amt']) + '<', text, count=1)
    if t2.get('iss_price'):
        text = re.sub(r'>원<', '>' + _xml_safe(t2['iss_price']) + '<', text, count=1)
    if t2.get('ratio'):
        text = re.sub(r'>%<', '>' + _xml_safe(t2['ratio']) + '<', text, count=1)

    # 4. 표4 적(Y)/부(N) 치환
    # 같은 셀 내에 적(Y)와 부(N)이 별도 <hp:p> 태그에 있음
    # 선택된 값 → 적색(charPrIDRef=156), 비선택 값 → 텍스트 제거
    for yn_val in yn_markers:
        if '적' in yn_val:
            # 적(Y)를 적색으로 표시 (기울임 없음)
            text = text.replace(
                'charPrIDRef="52"><hp:t>적(Y)</hp:t>',
                f'charPrIDRef="{RED_CHARPR_ID}"><hp:t>적(Y)</hp:t>',
                1
            )
            # 부(N) 텍스트 제거
            text = text.replace('>부(N)</hp:t>', '></hp:t>', 1)
        else:
            # 부(N)를 적색으로 표시 (기울임 없음)
            text = text.replace(
                'charPrIDRef="52"><hp:t>부(N)</hp:t>',
                f'charPrIDRef="{RED_CHARPR_ID}"><hp:t>부(N)</hp:t>',
                1
            )
            # 적(Y) 텍스트 제거
            text = text.replace('>적(Y)</hp:t>', '></hp:t>', 1)

    # 4.5 투자방법 (O) 체크 - "(   )" 를 "(O)" 또는 유지
    method_checks = replacements.get('_invest_method_checks', [])
    for i, checked in enumerate(method_checks):
        if checked:
            text = text.replace('(   )', '(O)', 1)
        else:
            # 체크 안된 항목도 한번 건너뜀 (순서 유지)
            idx = text.find('(   )')
            if idx >= 0:
                pass  # 그대로 유지
            # 다음 (   )로 이동하기 위해 임시 마킹 후 복원
            text = text.replace('(   )', '(___SKIP___)', 1)
    # 복원
    text = text.replace('(___SKIP___)', '(   )')

    # 5. 표5 실제 칼럼 빈 셀에 적/부 값 채우기
    table5_yn = replacements.get('_table5_yn', [])
    if table5_yn:
        t5_start = text.find('5. 준법사항 확인')
        if t5_start >= 0:
            before = text[:t5_start]
            after = text[t5_start:]

            # tc 단위로 분리하여 colAddr="2"인 빈 셀만 찾기
            tc_pattern = re.compile(r'(<hp:tc\b[^>]*>)(.*?)(</hp:tc>)', re.DOTALL)
            yn_idx = 0

            def _fill_cell(m):
                nonlocal yn_idx
                tc_open = m.group(1)
                tc_body = m.group(2)
                tc_close = m.group(3)

                # colAddr="2" 이고 빈 셀 (자체닫기 run)인 경우만
                if 'colAddr="2"' not in tc_body:
                    return m.group(0)
                row_m = re.search(r'rowAddr="(\d+)"', tc_body)
                row = int(row_m.group(1)) if row_m else -1
                if row < 2:
                    return m.group(0)

                # 자체닫기 run이 있는지 확인
                empty_run = re.search(r'<hp:run charPrIDRef="(\d+)"/>', tc_body)
                if not empty_run:
                    return m.group(0)

                # 이미 <hp:t>가 있으면 건너뜀
                if '<hp:t>' in tc_body:
                    return m.group(0)

                if yn_idx >= len(table5_yn):
                    return m.group(0)

                val = table5_yn[yn_idx]
                yn_idx += 1

                old_run = f'<hp:run charPrIDRef="{empty_run.group(1)}"/>'
                new_run = f'<hp:run charPrIDRef="{empty_run.group(1)}"><hp:t>{val}</hp:t></hp:run>'
                tc_body = tc_body.replace(old_run, new_run, 1)

                return tc_open + tc_body + tc_close

            after = tc_pattern.sub(_fill_cell, after)
            text = before + after

    # 6. 비고란(colAddr=3) 빈 셀에 적색 주석 채우기
    bigo_notes = replacements.get('_bigo_notes', [])
    if bigo_notes:
        t5_start = text.find('5. 준법사항 확인')
        if t5_start >= 0:
            before5 = text[:t5_start]
            after5 = text[t5_start:]

            tc_pattern = re.compile(r'(<hp:tc\b[^>]*>)(.*?)(</hp:tc>)', re.DOTALL)
            bigo_idx = 0

            def _fill_bigo(m):
                nonlocal bigo_idx
                tc_open = m.group(1)
                tc_body = m.group(2)
                tc_close = m.group(3)

                if 'colAddr="3"' not in tc_body:
                    return m.group(0)
                row_m = re.search(r'rowAddr="(\d+)"', tc_body)
                row = int(row_m.group(1)) if row_m else -1
                if row < 2:
                    return m.group(0)

                empty_run = re.search(r'<hp:run charPrIDRef="(\d+)"/>', tc_body)
                if not empty_run:
                    return m.group(0)
                if '<hp:t>' in tc_body:
                    return m.group(0)

                if bigo_idx >= len(bigo_notes):
                    return m.group(0)

                note = bigo_notes[bigo_idx]
                bigo_idx += 1

                if note:
                    safe_note = _xml_safe(note)
                    old_run = f'<hp:run charPrIDRef="{empty_run.group(1)}"/>'
                    # 적색+기울임 charPr(id=156) 사용
                    new_run = f'<hp:run charPrIDRef="{RED_ITALIC_CHARPR_ID}"><hp:t>{safe_note}</hp:t></hp:run>'
                    tc_body = tc_body.replace(old_run, new_run, 1)

                return tc_open + tc_body + tc_close

            after5 = tc_pattern.sub(_fill_bigo, after5)
            text = before5 + after5

    # 7. 담당자 확인 필요 행 → 행 전체를 적색(157)으로 변경
    # 역순으로 처리하여 앞쪽 위치에 영향 주지 않음
    red_full_rows = replacements.get('_red_full_rows', [])
    # 키워드 위치로 정렬 후 역순 처리
    red_positions = []
    for keyword in red_full_rows:
        idx = text.find(keyword)
        if idx >= 0:
            red_positions.append((idx, keyword))
    red_positions.sort(reverse=True)  # 뒤에서부터 처리

    for idx, keyword in red_positions:
        tr_start = text.rfind('<hp:tr', 0, idx)
        tr_end = text.find('</hp:tr>', idx)
        if tr_start >= 0 and tr_end >= 0:
            tr_end += len('</hp:tr>')
            tr_chunk = text[tr_start:tr_end]
            modified_tr = re.sub(r'charPrIDRef="\d+"', f'charPrIDRef="{RED_CHARPR_ID}"', tr_chunk)
            text = text[:tr_start] + modified_tr + text[tr_end:]

    # 8. 텍스트 기반 주석 (발굴경위, TCB등급, 투심위예정일)
    red_notes = replacements.get('_red_notes', {})
    for keyword, note in red_notes.items():
        if keyword in text:
            safe_note = _xml_safe(note)
            # keyword를 note로 완전 교체 (중복 방지)
            text = text.replace(keyword, safe_note, 1)

    return text


def _copy_and_replace(template_path: str, output_path: str, replacements: dict):
    """양식 HWPX를 복사하고 치환. flag_bits 등 원본 메타데이터 완전 보존."""
    import struct
    import shutil

    temp_path = output_path + '.tmp'

    with zipfile.ZipFile(template_path, 'r') as zin:
        modified = {}

        # section0.xml, PrvText.txt 치환
        for fname in ('Contents/section0.xml', 'Preview/PrvText.txt'):
            text = zin.read(fname).decode('utf-8', errors='replace')
            text = _apply_replacements(text, replacements)
            modified[fname] = text.encode('utf-8')

        # header.xml에 적색+기울임 charPr 추가 (id=156)
        header = zin.read('Contents/header.xml').decode('utf-8', errors='replace')
        header = _add_red_italic_charpr(header)
        modified['Contents/header.xml'] = header.encode('utf-8')

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

RED_ITALIC_CHARPR_ID = "156"  # 적색+기울임 (주석용)
RED_CHARPR_ID = "157"         # 적색만 (행 전체 강조용)


def _add_red_italic_charpr(header_xml: str) -> str:
    """header.xml에 적색+기울임(156)과 적색만(157) charPr을 추가."""
    import re
    m = re.search(r'(<hh:charPr\s+id="22".*?</hh:charPr>)', header_xml, re.DOTALL)
    if not m:
        return header_xml

    base = m.group(1)
    new_entries = ""

    # charPr 156: 적색 + 기울임 (주석용)
    if f'id="{RED_ITALIC_CHARPR_ID}"' not in header_xml:
        cp156 = base.replace('id="22"', f'id="{RED_ITALIC_CHARPR_ID}"')
        cp156 = cp156.replace('textColor="#000000"', 'textColor="#FF0000"')
        cp156 = cp156.replace('useFontSpace=', 'italic="1" useFontSpace=')
        new_entries += cp156 + "\n"

    # charPr 157: 적색만 (행 전체 강조용)
    if f'id="{RED_CHARPR_ID}"' not in header_xml:
        cp157 = base.replace('id="22"', f'id="{RED_CHARPR_ID}"')
        cp157 = cp157.replace('textColor="#000000"', 'textColor="#FF0000"')
        new_entries += cp157 + "\n"

    if new_entries:
        header_xml = header_xml.replace('</hh:charProperties>', new_entries + '</hh:charProperties>')
    return header_xml


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
