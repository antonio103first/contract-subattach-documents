---
description: 투자계약서와 투심보고서를 입력받아 체크리스트 2종(DOCX, HWPX)을 자동 생성합니다.
---

# generate-checklist

투자계약 전 작성해야 하는 2개의 체크리스트를 자동 생성하는 스킬입니다.

## 생성 문서
1. **투자계약서 체크리스트 및 의무기재사항확인서** (.docx)
2. **준법사항체크리스트** (.hwpx)

## 사용법
```
/generate-checklist [투자계약서파일] [투심보고서파일]
```

## 실행 방법
사용자가 투자계약서와 투심보고서 파일 경로를 제공하면, 다음 명령을 실행하세요:

```bash
cd /home/user/contract-subattach-documents && python generate_checklists.py \
  --contract "<투자계약서 파일 경로>" \
  --report "<투심보고서 파일 경로>" \
  --output-dir ./output
```

## 주의사항
- 투자계약서와 투심보고서는 `.docx` 형식이어야 합니다.
- 두 문서에서 추출한 데이터가 불일치할 경우 주석(Comment)으로 경고가 표시됩니다.
- 조문 번호는 키워드 기반으로 자동 탐색되므로 계약서 구조가 달라도 대응 가능합니다.
- 출력 파일은 `./output` 디렉토리에 저장됩니다.
