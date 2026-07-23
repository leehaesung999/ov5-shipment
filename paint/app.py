# -*- coding: utf-8 -*-
"""색칠 (창고 레이아웃) — 원본 로케이션 데이터를 레이아웃 xlsx에 채우고 색·테두리 적용."""
import sys
import traceback
from pathlib import Path
from datetime import datetime

import streamlit as st

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
import fill_locations as fl  # noqa: E402

TMP = HERE / "_tmp"
TMP.mkdir(exist_ok=True)


st.title("🎨 창고 레이아웃 색칠")
st.caption(
    "원본 로케이션 데이터(예: `0507 기본 로케이션.xlsx`)를 색칠 레이아웃 파일에 자동으로 채우고 "
    "네이비 헤더 · 노랑/초록 랙 · 회색 통로 등 시각 디자인을 적용합니다."
)

col1, col2 = st.columns(2)
with col1:
    up_src = st.file_uploader(
        "① 원본 데이터 xlsx",
        type=["xlsx"],
        help="로케이션ID → 품목 매핑이 담긴 원본 파일 (예: 기본 로케이션 xlsx)",
    )
with col2:
    up_layout = st.file_uploader(
        "② 레이아웃 파일 xlsx / xlsm",
        type=["xlsx", "xlsm"],
        help="비어있는 창고 배치도 (색칠할 대상)",
    )

with st.expander("옵션", expanded=False):
    sheet_name = st.text_input(
        "작업할 시트 이름 (비우면 '재고조사' 또는 첫 시트)",
        value="",
    )
    do_style = st.checkbox("시각 디자인 적용", value=True)

run_clicked = st.button("▶ 색칠 실행", type="primary",
                        disabled=not (up_src and up_layout),
                        use_container_width=True)

if not (up_src and up_layout):
    st.info("원본 데이터와 레이아웃 파일을 둘 다 업로드하세요.")

if run_clicked:
    try:
        src_path = TMP / f"_src_{up_src.name}"
        src_path.write_bytes(up_src.getvalue())
        layout_path = TMP / f"_layout_{up_layout.name}"
        layout_path.write_bytes(up_layout.getvalue())

        ts = datetime.now().strftime("%Y%m%d_%H%M")
        stem = Path(up_layout.name).stem
        suffix = Path(up_layout.name).suffix
        out_path = TMP / f"{stem}_완성_{ts}{suffix}"

        with st.spinner("색칠 진행 중..."):
            fl.run(
                src_path, layout_path, out_path,
                do_style=do_style,
                sheet_name=(sheet_name.strip() or None),
            )

        st.success(f"완료: {out_path.name}")
        st.download_button(
            "📥 결과 파일 다운로드",
            data=out_path.read_bytes(),
            file_name=out_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"오류: {e}")
        st.code(traceback.format_exc())

st.divider()
with st.expander("도움말"):
    st.markdown("""
- **원본 데이터**: 첫 열에 로케이션ID(`A11-01-10` 형태), 그다음 열에 품목코드·품목명이 있는 엑셀
- **레이아웃 파일**: 창고 배치도 형태의 엑셀 (색칠할 빈 시트)
  - 로케이션ID 셀 아래/옆에 품목이 자동 채워짐
  - 헤더(`A11`), 통로 행도 자동 인식
- **시각 디자인 적용 해제**: 값만 채우고 색은 안 넣음
- 시트 이름은 기본적으로 `재고조사`를 찾고, 없으면 첫 시트 사용
""")
