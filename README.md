# Excel 시트 번역기 (Excel Sheet Translator)

엑셀 파일의 한글 내용을 다른 언어로 번역해주는 도구입니다. 여러 시트가 있는 파일도 모두 처리하며, 용어 사전 기능으로 번역 일관성을 유지합니다.

## 주요 기능

- 엑셀 파일의 모든 시트에서 한글 셀 자동 감지
- GPT-4.1 모델을 활용한 고품질 번역
- 용어 사전(grocery) 기능으로 번역 일관성 유지
- 같은 시트 내에서도 동일 용어는 일관되게 번역
- 배치 처리로 대용량 시트도 안정적 처리
- 번역 맵 JSON 파일 생성으로 결과 검수 용이
- 상세 로깅 기능

## 사용 방법

1. 필요한 패키지 설치:
```
pip install openai openpyxl tqdm
```

2. 스크립트 실행:
```
python translator_new.py
```

3. 프롬프트에 따라 입력:
   - OpenAI API 키
   - 번역할 엑셀 파일 경로 
   - 원본 언어 코드 (예: Korean)
   - 대상 언어 코드 (예: English)

4. 결과 파일:
   - 번역된 엑셀 파일 (원본파일명_batched_translated_언어.xlsx)
   - 번역 맵 JSON 파일 (원본파일명_시트명_translation_map.json)
   - 용어 사전 파일 (원본파일명_translation_grocery.json)
   - 로그 파일 (translation_log.txt)

## 주의사항

- OpenAI API 키가 필요합니다
- 대용량 파일 번역 시 API 비용이 발생할 수 있습니다
- 용어 사전 기능으로 API 호출을 최소화합니다 (비용 절감)

## 라이센스

MIT License 