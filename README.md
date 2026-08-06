# 공구 매출 정리 도구

스마트스토어 상품별 매출 파일을 업로드하면 아래 네 열만 자동으로 정리하는 Streamlit 앱입니다.

- 벤더사
- 셀러
- 품목
- 판매금액

판매금액은 원본 파일의 **정상금액**을 사용합니다.

## 주요 기능

- XLSX, XLS, CSV 업로드
- 파일 위쪽의 빈 행과 `(부분취소)` 보조 행 자동 제외
- `김희경 x 무아스 2 in 1 스윙 핸디 스팀다리미옵션별 보기` 형식 자동 분리
- 셀러별 벤더사 자동 연결
- 미등록 벤더사 화면에서 직접 입력
- 동일 벤더사·셀러·품목 합산
- 정리 결과 XLSX 다운로드
- 자사상품 및 파싱되지 않은 상품 별도 확인

## 폴더 구성

```text
mooas_sales_converter/
├─ app.py
├─ requirements.txt
├─ vendor_mapping.csv
└─ README.md
```

## 내 컴퓨터에서 실행

Python을 설치한 뒤 이 폴더에서 터미널을 열고 실행합니다.

```bash
pip install -r requirements.txt
streamlit run app.py
```

브라우저에서 보통 아래 주소로 열립니다.

```text
http://localhost:8501
```

## Streamlit Community Cloud에 배포

1. GitHub 저장소를 만듭니다.
2. 이 폴더 안의 파일 4개를 저장소에 업로드합니다.
3. Streamlit Community Cloud에서 GitHub 저장소를 연결합니다.
4. 실행 파일로 `app.py`를 선택하고 배포합니다.

## 벤더 매핑 추가

`vendor_mapping.csv` 파일을 엑셀로 열어 아래 형식으로 셀러를 계속 추가하면 됩니다.

```csv
셀러,벤더사
김희경,트리에티
현맘,니드유
```

앱 화면의 `현재 벤더 매핑표 다운로드` 버튼으로 수정본을 받은 뒤,
GitHub의 `vendor_mapping.csv`를 해당 파일로 교체해도 됩니다.

## 원본 파일 필수 열

- 상품코드
- 상품명
- 정상금액

상품명은 아래 형식일 때 자동으로 셀러와 품목을 나눕니다.

```text
셀러명 x 무아스 품목명옵션별 보기
```
