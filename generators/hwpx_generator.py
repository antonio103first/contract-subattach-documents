"""준법사항체크리스트 HWPX 파일 생성 모듈.
실제 양식 HWPX를 복사하여 section0.xml의 placeholder를 치환하는 방식.
이렇게 하면 한컴오피스에서 정상적으로 열리는 유효한 HWPX를 생성한다."""
import os
import re
import shutil
import zipfile
from datetime import datetime


# 양식 HWPX 파일 경로
DEFAULT_TEMPLATE = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "(양식) 준법사항체크리스트(안)_2024 IBK혁신 케이런 모빌리티 7호 펀드_251208.hwpx"
)


def generate_hwpx_checklist(contract_data, report_data, output_path: str,
                             template_path: str = None):
    """양식 HWPX를 복사하여 placeholder를 치환하고 준법사항체크리스트를 생성한다."""
    template_path = template_path or DEFAULT_TEMPLATE

    if not os.path.exists(template_path):
        print(f"[ERROR] HWPX 양식 파일을 찾을 수 없습니다: {template_path}")
        print("[SKIP] HWPX 생성을 건너뜁니다.")
        return

    # 데이터 준비
    replacements = _build_replacements(contract_data, report_data)
    warnings = _check_mismatches(contract_data, report_data)

    # 양식 복사 후 section0.xml 치환
    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    _copy_and_replace(template_path, output_path, replacements)

    print(f"[OK] HWPX 준법사항체크리스트 생성 완료: {output_path}")

    if warnings:
        print(f"\n[주의] HWPX {len(warnings)}건의 불일치 발견:")
        for note in warnings:
            print(f"  - {note}")


def _build_replacements(contract_data, report_data) -> dict:
    """placeholder → 실제 값 매핑을 생성한다."""
    cd = contract_data
    rd = report_data

    # 회사명 (짧은 형태)
    company_name = rd.company_name or cd.company_name
    short_name = company_name
    if '주식회사' in company_name:
        short_name = "㈜" + company_name.replace('주식회사', '').replace('㈜', '').strip()
    if not short_name.startswith('㈜') and '㈜' not in short_name and '(주)' not in short_name:
        short_name = "㈜" + short_name

    representative = rd.representative or cd.representative or ""
    address = rd.address or cd.address or ""
    business_id = rd.business_registration or ""
    stock_type = cd.stock_type or rd.stock_type or ""

    investment_amount = cd.total_investment or ""
    if investment_amount and not investment_amount.endswith("원"):
        investment_amount += "원"
    issue_price = cd.issue_price or ""
    if issue_price and not issue_price.endswith("원"):
        issue_price += "원"
    share_ratio = rd.share_ratio or ""
    if share_ratio and not share_ratio.endswith("%"):
        share_ratio += "%"

    # 주요 투자조건
    conditions_parts = []
    if cd.duration:
        conditions_parts.append(f" - 존속기간 : {cd.duration}")
    if cd.redemption_terms:
        conditions_parts.append(f" - 상환조건 : {cd.redemption_terms}")
    if cd.conversion_terms:
        conditions_parts.append(f" - 전환조건 : {cd.conversion_terms}")
    if cd.refixing_terms:
        conditions_parts.append(f" - {cd.refixing_terms}")
    if cd.other_terms:
        conditions_parts.append(f" - 기타 : {cd.other_terms} 등")
    conditions_text = " ".join(conditions_parts) if conditions_parts else ""

    # 위약벌
    penalty_parts = []
    if cd.penalty_rate:
        penalty_parts.append(f" - 위약벌 : 투자금의 {cd.penalty_rate}%")
    if cd.delay_rate:
        penalty_parts.append(f" - 지연배상금 : 실제 지급일까지 연 {cd.delay_rate}%")
    if cd.buyback_rate:
        penalty_parts.append(f" - 주식매수청구권 : 투자원금 및 {cd.buyback_rate}%")
    penalty_text = " ".join(penalty_parts)

    discoverer = rd.discoverer or ""
    reviewer = rd.reviewer or ""
    post_manager = rd.post_manager or ""
    establishment_date = rd.establishment_date or ""

    # section0.xml 내 텍스트 치환 맵
    # 단순 replace 가능한 것들 (유일한 placeholder)
    simple = {
        '㈜AAA': short_name,
        '000-00-00000': business_id or '(미확인)',
    }

    # 순서 기반 치환이 필요한 것들
    # OOO가 4번 나옴: 대표이사, 발굴자, 심사자, 사후관리자
    # OO가 1번: 소재지
    ordered = {
        'OOO': [representative, discoverer, reviewer, post_manager],
        'OO': [address],
    }

    # 조건/위약벌 텍스트 치환
    condition_replacements = {}
    if conditions_text:
        condition_replacements[' - 존속기간 :  - 상환조건 :  - 전환조건 :  - 기타 :'] = conditions_text
    if penalty_text:
        condition_replacements[' - 위약벌 :  - 지연배상금 :  - 주식매수청구권 :'] = penalty_text

    return {
        '_simple': simple,
        '_ordered': ordered,
        '_conditions': condition_replacements,
        '_investment': {
            'stock_type': stock_type,
            'investment_amount': investment_amount,
            'issue_price': issue_price,
            'share_ratio': share_ratio,
        },
    }


def _check_mismatches(contract_data, report_data) -> list:
    """불일치 확인."""
    warnings = []
    cd = contract_data
    rd = report_data

    def _norm(v):
        return re.sub(r'[\s,원주%㈜주식회사(주)]', '', str(v or ''))

    checks = [
        ("투자금액", rd.investment_amount, (cd.total_investment + "원") if cd.total_investment else ""),
        ("투자단가", rd.issue_price, (cd.issue_price + "원") if cd.issue_price else ""),
        ("투자방식", rd.stock_type, cd.stock_type),
    ]
    for name, rv, cv in checks:
        if rv and cv and _norm(rv) != _norm(cv):
            warnings.append(f"{name}: 투심보고서={rv}, 투자계약서={cv}")
            print(f"[WARNING] {name} 불일치: 투심보고서={rv}, 투자계약서={cv}")

    return warnings


def _apply_replacements(text: str, replacements: dict) -> str:
    """모든 치환 규칙을 텍스트에 적용한다."""
    simple = replacements.get('_simple', {})
    ordered = replacements.get('_ordered', {})
    conditions = replacements.get('_conditions', {})

    # 1. 단순 치환 (유일한 placeholder)
    for old_val, new_val in simple.items():
        if old_val and new_val:
            text = text.replace(old_val, new_val)

    # 2. 조건/위약벌 텍스트 치환
    for old_val, new_val in conditions.items():
        if old_val and new_val:
            text = text.replace(old_val, new_val)

    # 3. 순서 기반 치환 (OOO가 여러 번 나오는 경우)
    for placeholder, values in ordered.items():
        for new_val in values:
            if new_val:
                # XML: >OOO< 패턴
                pat_xml = '>' + re.escape(placeholder) + '<'
                repl_xml = '>' + new_val + '<'
                text = re.sub(pat_xml, repl_xml, text, count=1)
                # PrvText: <OOO> 패턴 (꺽쇠가 구분자)
                pat_prv = '<' + re.escape(placeholder) + '>'
                repl_prv = '<' + new_val + '>'
                text = re.sub(pat_prv, repl_prv, text, count=1)

    return text


def _copy_and_replace(template_path: str, output_path: str, replacements: dict):
    """양식 HWPX를 복사하고 section0.xml의 placeholder를 치환한다."""
    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename in ('Contents/section0.xml', 'Preview/PrvText.txt'):
                    text = data.decode('utf-8', errors='replace')
                    text = _apply_replacements(text, replacements)
                    data = text.encode('utf-8')

                # mimetype은 압축하지 않음
                if item.filename == 'mimetype':
                    zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                else:
                    zout.writestr(item, data)
