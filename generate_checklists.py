#!/usr/bin/env python3
"""
투자 체크리스트 자동 생성 시스템

투자계약서와 투자심사보고서를 입력받아 다음 2종의 체크리스트를 자동 생성합니다:
1. 투자계약서 체크리스트 및 의무기재사항확인서 (.docx)
2. 준법사항체크리스트 (.hwpx)

사용법:
  python generate_checklists.py \\
    --contract "투자계약서.docx" \\
    --report "투심보고서.docx" \\
    --output-dir ./output
"""
import argparse
import os
import sys

from extractors.contract_extractor import extract_contract_data
from extractors.report_extractor import extract_report_data
from generators.docx_generator import generate_docx_checklist
from generators.hwpx_generator import generate_hwpx_checklist


# 기본 템플릿 경로
DEFAULT_TEMPLATE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DOCX_TEMPLATE = os.path.join(
    DEFAULT_TEMPLATE_DIR,
    "(양식) 투자계약서 체크리스트 및 의무기재사항확인서_2024 IBK 혁신펀드.docx"
)


def main():
    parser = argparse.ArgumentParser(
        description="투자 체크리스트 자동 생성 시스템"
    )
    parser.add_argument(
        "--contract", required=True,
        help="투자계약서 파일 경로 (.docx 또는 .pdf)"
    )
    parser.add_argument(
        "--report", required=True,
        help="투자심사보고서 파일 경로 (.docx 또는 .pdf)"
    )
    parser.add_argument(
        "--output-dir", default="./output",
        help="출력 디렉토리 (기본: ./output)"
    )
    parser.add_argument(
        "--docx-template", default=DEFAULT_DOCX_TEMPLATE,
        help="투자계약서 체크리스트 DOCX 양식 파일 경로"
    )
    args = parser.parse_args()

    # 파일 존재 확인
    for path, name in [(args.contract, "투자계약서"), (args.report, "투심보고서")]:
        if not os.path.exists(path):
            print(f"[ERROR] {name} 파일을 찾을 수 없습니다: {path}")
            sys.exit(1)

    if not os.path.exists(args.docx_template):
        print(f"[ERROR] DOCX 양식 파일을 찾을 수 없습니다: {args.docx_template}")
        sys.exit(1)

    os.makedirs(args.output_dir, exist_ok=True)

    # 1단계: 투자계약서 데이터 추출
    print("=" * 60)
    print("[1/4] 투자계약서 데이터 추출 중...")
    contract_data = extract_contract_data(args.contract)
    print(f"  회사명: {contract_data.company_name}")
    print(f"  투자금액: {contract_data.total_investment}원")
    print(f"  투자단가: {contract_data.issue_price}원")
    print(f"  발행주식수: {contract_data.total_shares}주")
    print(f"  투자방식: {contract_data.stock_type}")
    print(f"  조문 - 투자금용도: 제{contract_data.article_fund_usage}조")
    print(f"  조문 - 주식매수청구권: 제{contract_data.article_buyback}조")
    print(f"  조문 - 손해배상/위약벌: 제{contract_data.article_damages}조")
    print(f"  조문 - 지연배상금: 제{contract_data.article_delay_penalty}조")

    # 2단계: 투심보고서 데이터 추출
    print()
    print("[2/4] 투자심사보고서 데이터 추출 중...")
    report_data = extract_report_data(args.report)
    print(f"  회사명: {report_data.company_name}")
    print(f"  대표이사: {report_data.representative}")
    print(f"  사업자번호: {report_data.business_registration}")
    print(f"  지분율: {report_data.share_ratio}")
    print(f"  펀드명: {report_data.fund_name}")
    print(f"  발굴자: {report_data.discoverer}")
    print(f"  심사자: {report_data.reviewer}")

    # 회사명으로 출력 파일명 구성 (㈜ 등 제거하고 본명만)
    from extractors.contract_extractor import normalize_company_name
    company_full = report_data.company_name or contract_data.company_name
    company_short = normalize_company_name(company_full).lstrip('㈜')

    # 3단계: DOCX 체크리스트 생성
    print()
    print("[3/4] 투자계약서 체크리스트 DOCX 생성 중...")
    docx_output = os.path.join(
        args.output_dir,
        f"투자계약서 체크리스트 및 의무기재사항확인서_{company_short}.docx"
    )
    warnings = generate_docx_checklist(
        contract_data, report_data, args.docx_template, docx_output
    )

    # 4단계: HWPX 준법사항체크리스트 생성
    print()
    print("[4/4] 준법사항체크리스트 HWPX 생성 중...")
    fund_short = report_data.fund_name or "펀드"
    hwpx_output = os.path.join(
        args.output_dir,
        f"준법사항체크리스트_{fund_short}_{company_short}.hwpx"
    )
    generate_hwpx_checklist(contract_data, report_data, hwpx_output)

    # 완료 요약
    print()
    print("=" * 60)
    print("생성 완료!")
    print(f"  DOCX: {docx_output}")
    print(f"  HWPX: {hwpx_output}")
    if warnings:
        print(f"\n  ⚠ {len(warnings)}건의 데이터 불일치가 발견되었습니다. 주석을 확인하세요.")
    print("=" * 60)


if __name__ == "__main__":
    main()
