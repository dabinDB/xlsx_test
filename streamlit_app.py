from datetime import date

import streamlit as st

from src.excel_template_tools import (
    build_data_preview,
    build_ai_prompt,
    compare_snapshots,
    count_injectable_values,
    extract_template_structure,
    generate_mapping_with_gemini,
    inject_data,
    mapping_as_text,
    normalize_daily_data,
    read_data_table,
    snapshot_design,
    validate_mapping,
)


APP_VERSION = "2026-05-11-multi-sheet-mapping-v1"


def get_secret(name: str) -> str:
    try:
        return st.secrets.get(name, "")
    except Exception:
        return ""

st.set_page_config(page_title="Excel Template Data Replacer", layout="wide")

st.title("Excel Template Data Replacer")
st.caption(
    "엑셀 템플릿의 서식과 수식은 유지하고, 업로드한 데이터 값만 교체하는 테스트 앱입니다. "
    f"App version: {APP_VERSION}"
)

template_file = st.file_uploader("1. 템플릿 엑셀 업로드", type=["xlsx"])
data_file = st.file_uploader("2. 교체 데이터 업로드", type=["csv", "xlsx"])

report_date = st.date_input("보고일자", value=date.today())
note = st.text_area("비고", placeholder="결과 파일의 note_cell에 들어갈 메모")

with st.sidebar:
    st.header("Gemini API")
    secret_api_key = get_secret("GEMINI_API_KEY")
    api_key_input = st.text_input(
        "Gemini API Key",
        type="password",
        value="" if secret_api_key else "",
        placeholder="Streamlit Secrets에 없을 때만 입력",
    )
    gemini_api_key = secret_api_key or api_key_input
    gemini_model = st.text_input("Model", value=get_secret("GEMINI_MODEL") or "gemini-2.5-flash")
    if secret_api_key:
        st.success("GEMINI_API_KEY가 Secrets에서 로드되었습니다.")

st.divider()

if "mapping_text" not in st.session_state:
    st.session_state.mapping_text = mapping_as_text()

if data_file:
    data_df = read_data_table(data_file)
else:
    data_df = None

left, right = st.columns([1, 1])

with left:
    st.subheader("템플릿 구조 / AI 프롬프트")
    if template_file:
        template_bytes = template_file.getvalue()
        structure = extract_template_structure(template_bytes)
        sheet_names = [sheet["name"] for sheet in structure["sheets"]]
        st.success(f"분석된 시트 수: {len(sheet_names)}개")
        st.write("시트 목록:", ", ".join(sheet_names))
        st.json(structure, expanded=False)
        data_preview = build_data_preview(data_df) if data_df is not None else None
        with st.expander("AI 매핑 프롬프트 보기"):
            st.code(build_ai_prompt(structure, data_preview=data_preview), language="text")
        if data_preview:
            st.write("업로드 데이터 컬럼:", ", ".join(data_preview["columns"]))
        if st.button("Gemini API로 템플릿+데이터 매핑 생성", disabled=not (gemini_api_key and data_preview), type="secondary"):
            try:
                with st.spinner("Gemini API가 템플릿 매핑을 추론하는 중입니다..."):
                    generated_mapping = generate_mapping_with_gemini(
                        structure=structure,
                        api_key=gemini_api_key,
                        model=gemini_model,
                        data_preview=data_preview,
                    )
                st.session_state.mapping_text = mapping_as_text(generated_mapping)
                st.success("매핑 JSON을 생성했습니다. 오른쪽 입력창을 확인하세요.")
            except Exception as exc:
                message = str(exc)
                if "API key" in message or "permission" in message.lower() or "unauthenticated" in message.lower():
                    st.error("Gemini API Key가 유효하지 않거나 권한이 없습니다. Streamlit Secrets 또는 사이드바 입력값을 확인하세요.")
                elif "quota" in message.lower() or "rate" in message.lower():
                    st.error("Gemini API 사용량 한도 또는 쿼터에 도달했습니다.")
                    st.info("Google AI Studio/Google Cloud의 API 키, 결제, 무료 한도, 분당 요청 제한을 확인하세요.")
                else:
                    st.error(f"Gemini API 호출 중 오류가 발생했습니다: {message}")
        if not gemini_api_key:
            st.info("Gemini API 매핑 생성을 쓰려면 사이드바에 API 키를 입력하거나 Streamlit Secrets를 설정하세요.")
        if data_preview is None:
            st.info("업로드 데이터 컬럼까지 AI로 매핑하려면 교체 데이터 파일도 업로드하세요.")
    else:
        template_bytes = None
        st.info("템플릿 파일을 업로드하면 구조 분석 결과와 AI 프롬프트가 표시됩니다.")

with right:
    st.subheader("매핑 JSON")
    mapping_text = st.text_area(
        "AI가 반환한 매핑 JSON 또는 수동 매핑을 입력하세요.",
        key="mapping_text",
        height=420,
    )

st.divider()

if data_df is not None:
    st.subheader("업로드 데이터 미리보기")
    st.dataframe(data_df, use_container_width=True)

run = st.button("값 교체 및 검증 실행", type="primary", disabled=not (template_file and data_file))

if run and template_bytes and data_df is not None:
    try:
        mapping = validate_mapping(mapping_text)
        daily_data = normalize_daily_data(data_df, str(report_date), note)
        coverage = count_injectable_values(mapping, daily_data)

        st.subheader("데이터 매핑 확인")
        st.json(
            {
                "mapped_sheets": [sheet.get("sheet_name") for sheet in mapping.get("sheets", [])],
                "injectable_values": coverage["injectable_values"],
                "sheet_counts": coverage["sheet_counts"],
                "missing_sources": coverage["missing_sources"],
            },
            expanded=False,
        )
        if coverage["injectable_values"] == 0:
            st.error("업로드 데이터에서 주입할 값을 찾지 못했습니다. 매핑 JSON의 source_columns를 확인하세요.")
            st.stop()

        before = snapshot_design(template_bytes)
        output_bytes = inject_data(template_bytes, mapping, daily_data)
        after = snapshot_design(output_bytes)
        diff = compare_snapshots(before, after)

        st.subheader("검증 결과")
        if not diff["structural_diffs"] and diff["style_diffs_count"] == 0:
            st.success("구조와 주요 셀 서식이 보존되었습니다.")
        else:
            st.warning("일부 차이가 발견되었습니다. 아래 내용을 확인하세요.")

        c1, c2, c3 = st.columns(3)
        total_merged_ranges = sum(len(sheet["merged_ranges"]) for sheet in after["sheets"].values())
        total_charts = sum(sheet["chart_count"] for sheet in after["sheets"].values())
        c1.metric("전체 병합 영역", total_merged_ranges)
        c2.metric("전체 차트 수", total_charts)
        c3.metric("서식 차이", diff["style_diffs_count"])

        if diff["structural_diffs"]:
            st.error("\n".join(diff["structural_diffs"]))
        if diff["style_diffs_count"]:
            st.dataframe(diff["style_diffs_sample"], use_container_width=True)

        st.download_button(
            "결과 엑셀 다운로드",
            data=output_bytes,
            file_name=f"report_{report_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error(str(exc))
        st.info("매핑 JSON 입력창이 비어 있으면 먼저 `Gemini API로 템플릿+데이터 매핑 생성`을 누르거나 JSON을 직접 붙여넣어 주세요.")
