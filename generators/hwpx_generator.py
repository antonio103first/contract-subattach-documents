"""준법사항체크리스트 HWPX 파일 생성 모듈. ZIP+XML 직접 생성."""
import os
import zipfile
from datetime import datetime
from xml.sax.saxutils import escape


def generate_hwpx_checklist(contract_data, report_data, output_path: str):
    """준법사항체크리스트 HWPX 파일을 생성한다."""
    now = datetime.now()
    date_str = f"{now.year}년 {now.month}월 {now.day}일"

    # 데이터 준비
    company_name = report_data.company_name or contract_data.company_name
    short_name = company_name
    if '주식회사' in company_name:
        short_name = "㈜" + company_name.replace('주식회사', '').replace('㈜', '').strip()
    if not short_name.startswith('㈜') and '㈜' not in short_name:
        short_name = "㈜" + short_name

    representative = report_data.representative or contract_data.representative
    address = report_data.address or contract_data.address
    business_id = report_data.business_registration or ""
    fund_name = report_data.fund_name or "2024 IBK혁신 케이런 모빌리티 7호 펀드"

    investment_amount = contract_data.total_investment or ""
    if investment_amount:
        investment_amount = f"{investment_amount}원"
    issue_price = contract_data.issue_price or ""
    if issue_price:
        issue_price = f"{issue_price}원"
    share_ratio = report_data.share_ratio or ""
    stock_type = contract_data.stock_type or report_data.stock_type or "상환전환우선주"

    # 주요 투자조건
    conditions_parts = []
    if contract_data.duration:
        conditions_parts.append(f" - 존속기간 : {contract_data.duration}")
    if contract_data.redemption_terms:
        conditions_parts.append(f" - 상환조건 : {contract_data.redemption_terms}")
    if contract_data.conversion_terms:
        conditions_parts.append(f" - 전환조건 : {contract_data.conversion_terms}")
    if contract_data.refixing_terms:
        conditions_parts.append(f" - {contract_data.refixing_terms}")
    if contract_data.other_terms:
        conditions_parts.append(f" - 기타 : {contract_data.other_terms} 등")
    conditions_text = "\n".join(conditions_parts) if conditions_parts else ""

    # 위약벌
    penalty_parts = []
    if contract_data.penalty_rate:
        penalty_parts.append(f" - 위약벌 : 투자금의 {contract_data.penalty_rate}%")
    if contract_data.delay_rate:
        penalty_parts.append(f" - 지연배상금 : 실제 지급일까지 연 {contract_data.delay_rate}%")
    if contract_data.buyback_rate:
        penalty_parts.append(f" - 주식매수청구권 : 투자원금 및 {contract_data.buyback_rate}%")
    penalty_text = "\n".join(penalty_parts)

    # 담당자
    discoverer = report_data.discoverer or ""
    reviewer = report_data.reviewer or ""
    post_manager = report_data.post_manager or ""

    # 동반투자
    co_investor_rows = ""
    for name, amount, price in report_data.co_investors:
        co_investor_rows += f"<{escape(name)}><{escape(amount)}><{escape(price or issue_price)}>"

    # 설립일
    establishment_date = report_data.establishment_date or ""

    # 불일치 검사
    mismatch_notes = []
    _check_mismatch(mismatch_notes, "투자금액", report_data.investment_amount,
                    (contract_data.total_investment + "원") if contract_data.total_investment else "")
    _check_mismatch(mismatch_notes, "투자단가", report_data.issue_price,
                    (contract_data.issue_price + "원") if contract_data.issue_price else "")
    _check_mismatch(mismatch_notes, "투자방식", report_data.stock_type, contract_data.stock_type)

    # HWPX 파일 생성
    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)

    section_xml = _build_section_xml(
        fund_name=fund_name,
        short_name=short_name,
        representative=representative,
        business_id=business_id,
        address=address,
        stock_type=stock_type,
        investment_amount=investment_amount,
        issue_price=issue_price,
        share_ratio=share_ratio,
        conditions_text=conditions_text,
        penalty_text=penalty_text,
        discoverer=discoverer,
        reviewer=reviewer,
        post_manager=post_manager,
        establishment_date=establishment_date,
        co_investor_rows=co_investor_rows,
        date_str=date_str,
        mismatch_notes=mismatch_notes,
    )

    _write_hwpx(output_path, section_xml)
    print(f"[OK] HWPX 준법사항체크리스트 생성 완료: {output_path}")

    if mismatch_notes:
        print(f"\n[주의] HWPX {len(mismatch_notes)}건의 불일치 발견:")
        for note in mismatch_notes:
            print(f"  - {note}")


def _check_mismatch(notes: list, field_name: str, report_val: str, contract_val: str):
    """불일치 확인."""
    import re
    if not report_val or not contract_val:
        return
    rv = re.sub(r'[\s,원주%㈜주식회사(주)]', '', str(report_val))
    cv = re.sub(r'[\s,원주%㈜주식회사(주)]', '', str(contract_val))
    if rv != cv:
        notes.append(f"{field_name}: 투심보고서={report_val}, 투자계약서={contract_val}")
        print(f"[WARNING] {field_name} 불일치: 투심보고서={report_val}, 투자계약서={contract_val}")


def _build_section_xml(**d):
    """준법사항체크리스트 본문 XML을 생성한다."""
    e = escape

    xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
        xmlns:hc="http://www.hancom.co.kr/hwpml/2011/content"
        xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">

  {_p("＜준법사항 체크리스트(준법감시보고서)＞", bold=True, size=24)}

  {_p("")}
  {_p(f"▣ 펀드명 : {e(d['fund_name'])}")}
  {_p("▣ 업무집행조합원 : 케이런벤처스(유)")}
  {_p("")}

  {_p("1. 투자기업", bold=True)}
  {_table_company(d)}

  {_p("")}
  {_p("2. 투자내용", bold=True)}
  {_p("")}
  {_p("□ 투자유형 및 투자금액")}
  {_table_investment(d)}

  {_p("")}
  {_p(f"□ 주요 투자조건")}
  {_multiline_p(d['conditions_text'])}

  {_p("")}
  {_p("□ 위약벌 사항 등")}
  {_multiline_p(d['penalty_text'])}

  {_p("")}
  {_p("□ 동반투자내역(* 운용사내 타 펀드, 타 운용사 동반 투자내역)")}
  {_p(d.get('co_investor_rows', '(해당 내용 직접 확인 필요)'))}

  {_p("")}
  {_p("3. 투자 담당자", bold=True)}
  {_table_staff(d)}

  {_p("")}
  {_p("4. 투자기업의 벤처기업 등 해당여부 확인", bold=True)}
  {_p(f"  - 설립일: {e(d['establishment_date'])}")}
  {_p("  (벤처기업 확인서, 이노비즈 인증 등은 별도 확인 필요)")}

  {_p("")}
  {_p(f"작성일: {e(d['date_str'])}")}

  {_mismatch_section(d.get('mismatch_notes', []))}

</hs:sec>'''
    return xml


def _p(text: str, bold=False, size=20, color=None):
    """HWPX 문단 XML을 생성."""
    e = escape
    bold_attr = ' fontweight="bold"' if bold else ''
    color_attr = f' color="#{color}"' if color else ''
    return f'''<hp:p>
    <hp:run>
      <hp:rPr sz="{size}"{bold_attr}{color_attr}/>
      <hp:t>{e(text)}</hp:t>
    </hp:run>
  </hp:p>'''


def _multiline_p(text: str):
    """여러 줄 텍스트를 각각 문단으로."""
    if not text:
        return _p("")
    return "\n".join(_p(line) for line in text.split("\n"))


def _table_company(d):
    """투자기업 정보 테이블."""
    e = escape
    return f'''<hp:tbl>
    <hp:tr>
      <hp:tc><hp:p><hp:run><hp:t>업체명</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>{e(d['short_name'])}</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>대표이사</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>{e(d['representative'])}</hp:t></hp:run></hp:p></hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc><hp:p><hp:run><hp:t>사업자등록번호</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>{e(d['business_id'])}</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>소재지</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>{e(d['address'])}</hp:t></hp:run></hp:p></hp:tc>
    </hp:tr>
  </hp:tbl>'''


def _table_investment(d):
    """투자유형 및 투자금액 테이블."""
    e = escape
    return f'''<hp:tbl>
    <hp:tr>
      <hp:tc><hp:p><hp:run><hp:t>구 분</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>투자금액</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>투자단가</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>지분율</hp:t></hp:run></hp:p></hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc><hp:p><hp:run><hp:t>{e(d['stock_type'])}</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>{e(d['investment_amount'])}</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>{e(d['issue_price'])}</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>{e(d['share_ratio'])}</hp:t></hp:run></hp:p></hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc><hp:p><hp:run><hp:t>합 계</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>{e(d['investment_amount'])}</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>{e(d['issue_price'])}</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>{e(d['share_ratio'])}</hp:t></hp:run></hp:p></hp:tc>
    </hp:tr>
  </hp:tbl>'''


def _table_staff(d):
    """투자 담당자 테이블."""
    e = escape
    return f'''<hp:tbl>
    <hp:tr>
      <hp:tc><hp:p><hp:run><hp:t>발굴자 (기여율)</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>{e(d['discoverer'])}</hp:t></hp:run></hp:p></hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc><hp:p><hp:run><hp:t>심사자 (기여율)</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>{e(d['reviewer'])}</hp:t></hp:run></hp:p></hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc><hp:p><hp:run><hp:t>사후관리자 (기여율)</hp:t></hp:run></hp:p></hp:tc>
      <hp:tc><hp:p><hp:run><hp:t>{e(d['post_manager'])}</hp:t></hp:run></hp:p></hp:tc>
    </hp:tr>
  </hp:tbl>'''


def _mismatch_section(notes):
    """불일치 사항이 있을 경우 빨간색으로 표시."""
    if not notes:
        return ""
    lines = [_p(""), _p("[※ 투심보고서-계약서 불일치 항목]", bold=True, color="FF0000")]
    for note in notes:
        lines.append(_p(f"  - {note}", color="FF0000"))
    return "\n".join(lines)


def _write_hwpx(output_path: str, section_xml: str):
    """HWPX 파일(ZIP+XML)을 생성한다."""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        # mimetype (first entry, uncompressed)
        zf.writestr('mimetype', 'application/hwp+zip', compress_type=zipfile.ZIP_STORED)

        # META-INF/manifest.xml
        manifest = '''<?xml version="1.0" encoding="UTF-8"?>
<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <odf:file-entry odf:full-path="/" odf:media-type="application/hwp+zip"/>
  <odf:file-entry odf:full-path="Contents/content.hpf" odf:media-type="application/xml"/>
  <odf:file-entry odf:full-path="Contents/header.xml" odf:media-type="application/xml"/>
  <odf:file-entry odf:full-path="Contents/section0.xml" odf:media-type="application/xml"/>
</odf:manifest>'''
        zf.writestr('META-INF/manifest.xml', manifest)

        # Contents/content.hpf
        content_hpf = '''<?xml version="1.0" encoding="UTF-8"?>
<hp:HWPDocumentPackage xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
  xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
  version="1.1">
  <hp:compatibledocument target="HWP 2022"/>
</hp:HWPDocumentPackage>'''
        zf.writestr('Contents/content.hpf', content_hpf)

        # Contents/header.xml - A4 용지, 여백 설정
        header_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"
         xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
         xmlns:hc="http://www.hancom.co.kr/hwpml/2011/content">
  <hh:beginNum page="1" footnote="1" endnote="1"/>
  <hh:refList>
    <hh:fontfaces>
      <hh:fontface lang="HANGUL">
        <hh:font id="0" face="함초롬바탕" type="TTF"/>
      </hh:fontface>
      <hh:fontface lang="LATIN">
        <hh:font id="0" face="함초롬바탕" type="TTF"/>
      </hh:fontface>
    </hh:fontfaces>
    <hh:charProperties>
      <hh:charPr id="0" height="1000" color="0">
        <hh:fontRef hangul="0" latin="0"/>
      </hh:charPr>
    </hh:charProperties>
    <hh:paraProperties>
      <hh:paraPr id="0">
        <hh:align horizontal="JUSTIFY" vertical="BASELINE"/>
      </hh:paraPr>
    </hh:paraProperties>
  </hh:refList>
  <hh:secProperties>
    <hh:secPr>
      <hh:pageProperty paperWidth="59528" paperHeight="84188" landscape="NARROWLY">
        <hh:margin header="4252" footer="4252"
                   left="8504" right="8504" top="5668" bottom="4252" gutter="0"/>
      </hh:pageProperty>
    </hh:secPr>
  </hh:secProperties>
</hh:head>'''
        zf.writestr('Contents/header.xml', header_xml)

        # Contents/section0.xml - 본문
        zf.writestr('Contents/section0.xml', section_xml)

        # settings.xml
        settings = '''<?xml version="1.0" encoding="UTF-8"?>
<config:settings xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">
  <config:config-item config:name="ViewZoom" config:type="int">100</config:config-item>
</config:settings>'''
        zf.writestr('settings.xml', settings)
