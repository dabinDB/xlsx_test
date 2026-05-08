# Excel Template Data Replacer

엑셀 템플릿의 레이아웃, 서식, 수식, 차트는 유지하고 업로드한 일일 데이터 값만 교체하는 Streamlit 테스트 앱입니다.

## 기능

- 템플릿 엑셀 파일 업로드 (`.xlsx`)
- 템플릿 구조 추출 및 AI 매핑 프롬프트 생성
- 여러 시트가 있는 템플릿의 전체 시트 구조를 프롬프트에 포함
- CSV/XLSX 일일 데이터 업로드
- 매핑 JSON 기반으로 값만 교체
- 병합 셀, 컬럼 너비, 행 높이, 차트 수, 주요 셀 서식 보존 여부 검증
- 결과 엑셀 다운로드

## 로컬 실행

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Streamlit Cloud 배포

1. 이 폴더를 GitHub 저장소로 push합니다.
2. [Streamlit Community Cloud](https://streamlit.io/cloud)에서 새 앱을 만듭니다.
3. 저장소와 브랜치를 선택하고 main file path를 `streamlit_app.py`로 지정합니다.
4. Deploy를 누릅니다.

## 사용 흐름

1. 템플릿 엑셀 파일을 업로드합니다.
2. 앱이 보여주는 AI 프롬프트를 Claude, ChatGPT 등에게 보내 매핑 JSON을 얻습니다.
3. 매핑 JSON 입력창에 붙여넣습니다. 여러 시트 템플릿이면 `sheet_name`에 값을 교체할 시트명을 넣습니다.
4. `examples/sample_daily_data.csv` 형식처럼 교체 데이터를 업로드합니다.
5. 보고일자와 비고를 입력한 뒤 실행합니다.

기본 매핑 예시는 `examples/sample_mapping.json`에 있습니다.

현재 앱은 `.xlsx` 템플릿을 대상으로 합니다. 매크로가 있는 `.xlsm` 파일까지 보존하려면 `openpyxl.load_workbook(..., keep_vba=True)` 흐름으로 별도 확장이 필요합니다.
