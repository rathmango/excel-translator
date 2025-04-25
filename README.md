# Excel 시트 번역기 (Excel Sheet Translator)

엑셀 파일의 한글 내용을 다른 언어로 번역해주는 도구입니다. 여러 시트가 있는 파일도 모두 처리하며, 용어 사전 기능으로 번역 일관성을 유지합니다.

## 주요 기능

- 엑셀 파일(XLSX)과 CSV 파일 모두 지원
- 엑셀 파일의 모든 시트에서 한글 셀 자동 감지
- 100개 이상의 언어 간 번역 지원 (드롭다운 메뉴로 쉽게 선택)
- GPT-4.1 모델을 활용한 고품질 번역
- 용어 사전(grocery) 기능으로 번역 일관성 유지
- 같은 시트 내에서도 동일 용어는 일관되게 번역
- 배치 처리로 대용량 시트도 안정적 처리
- 번역 맵 JSON 파일 생성으로 결과 검수 용이
- 상세 로깅 기능

## 사용 방법

1. 필요한 패키지 설치:
```
pip install openai openpyxl pandas inquirer tqdm
```

2. 스크립트 실행:
```
python excel-translator.py
```

3. 프롬프트에 따라 입력:
   - OpenAI API 키
   - 번역할 파일 경로 (XLSX 또는 CSV)
   - 원본 언어 (드롭다운 메뉴에서 선택)
   - 대상 언어 (드롭다운 메뉴에서 선택)

4. 결과 파일:
   - 번역된 파일 (원본파일명_translated_언어코드.xlsx 또는 .csv)
   - 번역 맵 JSON 파일 (원본파일명_시트명_translation_map.json)
   - 용어 사전 파일 (원본파일명_translation_grocery.json)
   - 로그 파일 (translation_log.txt)

## 주의사항

- OpenAI API 키가 필요합니다
- 대용량 파일 번역 시 API 비용이 발생할 수 있습니다
- 용어 사전 기능으로 API 호출을 최소화합니다 (비용 절감)

## 라이센스

MIT License 