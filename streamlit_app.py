from datetime import date

import streamlit as st

from src.excel_template_tools import (
    build_ai_prompt,
    compare_snapshots,
    extract_template_structure,
    inject_data,
    mapping_as_text,
    normalize_daily_data,
    read_data_table,
    snapshot_design,
    validate_mapping,
)


st.set_page_config(page_title="Excel Template Data Replacer", layout="wide")

st.title("Excel Template Data Replacer")
st.caption("엑셀 템플릿의 서식과 수식은 유지하고, 업로드한 데이터 값만 교체하는 테스트 앱입니다.")

template_file = st.file_uploader("1. 템플릿 엑셀 업로드", type=["xlsx"])
data_file = st.file_uploader("2. 교체 데이터 업로드", type=["csv", "xlsx"])

report_date = st.date_input("보고일자", value=date.today())
note = st.text_area("비고", placeholder="결과 파일의 note_cell에 들어갈 메모")

st.divider()

left, right = st.columns([1, 1])

with left:
    st.subheader("템플릿 구조 / AI 프롬프트")
    if template_file:
        template_bytes = template_file.getvalue()
        structure = extract_template_structure(template_bytes)
        st.json(structure, expanded=False)
        with st.expander("AI 매핑 프롬프트 보기"):
            st.code(build_ai_prompt(structure), language="text")
    else:
        template_bytes = None
        st.info("템플릿 파일을 업로드하면 구조 분석 결과와 AI 프롬프트가 표시됩니다.")

with right:
    st.subheader("매핑 JSON")
    mapping_text = st.text_area(
        "AI가 반환한 매핑 JSON 또는 수동 매핑을 입력하세요.",
        value=mapping_as_text(),
        height=420,
    )

st.divider()

if data_file:
    data_df = read_data_table(data_file)
    st.subheader("업로드 데이터 미리보기")
    st.dataframe(data_df, use_container_width=True)
else:
    data_df = None

run = st.button("값 교체 및 검증 실행", type="primary", disabled=not (template_file and data_file))

if run and template_bytes and data_df is not None:
    try:
        mapping = validate_mapping(mapping_text)
        daily_data = normalize_daily_data(data_df, str(report_date), note)

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
        c1.metric("병합 영역", len(after["merged_ranges"]))
        c2.metric("차트 수", after["chart_count"])
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
        st.exception(exc)
