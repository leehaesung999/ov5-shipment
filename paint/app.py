# -*- coding: utf-8 -*-
"""색칠 (창고 레이아웃) — 원본 로케이션 데이터를 레이아웃 xlsm에 채우고 색·테두리 적용.

레이아웃 파일은 저장소에 내장(paint/data/기본레이아웃.xlsm) — 원본만 매번 업로드.
필요 시 옵션으로 사용자 지정 레이아웃 업로드 가능.
"""
import sys
import traceback
from pathlib import Path
from datetime import datetime

import streamlit as st

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
import fill_locations as fl  # noqa: E402

DATA = HERE / "data"
DEFAULT_LAYOUT = DATA / "기본레이아웃.xlsm"
TMP = HERE / "_tmp"
TMP.mkdir(exist_ok=True)


st.title("🎨 창고 레이아웃 색칠")
st.caption(
    "원본 로케이션 데이터를 창고 레이아웃에 자동으로 채우고 "
    "네이비 헤더 · 노랑/초록 랙 · 회색 통로 등 시각 디자인을 적용합니다."
)

up_src = st.file_uploader(
    "① 원본 데이터 xlsx",
    type=["xlsx"],
    help="로케이션ID → 품목 매핑이 담긴 원본 파일 (예: 기본 로케이션 xlsx)",
)

with st.expander("옵션 — 레이아웃 파일 · 시트 · 디자인", expanded=False):
    up_layout = st.file_uploader(
        "레이아웃 파일 xlsx / xlsm (선택, 미업로드 시 내장 기본 레이아웃 사용)",
        type=["xlsx", "xlsm"],
    )
    sheet_name = st.text_input(
        "작업할 시트 이름 (비우면 '재고조사' 또는 첫 시트)",
        value="",
    )
    do_style = st.checkbox("시각 디자인 적용", value=True)

# 레이아웃 상태 표시
if not DEFAULT_LAYOUT.exists() and not up_layout:
    st.warning("내장 레이아웃 파일이 없고 업로드도 안 되었습니다. 옵션에서 레이아웃 파일을 업로드하세요.")
    layout_ready = False
elif up_layout:
    st.caption(f"레이아웃: 업로드 파일 사용 (`{up_layout.name}`)")
    layout_ready = True
else:
    st.caption(f"레이아웃: 내장 기본 레이아웃 사용 (`{DEFAULT_LAYOUT.name}`)")
    layout_ready = True

run_clicked = st.button("▶ 색칠 실행", type="primary",
                        disabled=not (up_src and layout_ready),
                        width='stretch')

if not up_src:
    st.info("원본 데이터 xlsx를 업로드하세요.")

if run_clicked:
    try:
        src_path = TMP / f"_src_{up_src.name}"
        src_path.write_bytes(up_src.getvalue())

        if up_layout:
            layout_path = TMP / f"_layout_{up_layout.name}"
            layout_path.write_bytes(up_layout.getvalue())
        else:
            layout_path = DEFAULT_LAYOUT

        ts = datetime.now().strftime("%Y%m%d_%H%M")
        stem = Path(layout_path.name).stem
        suffix = Path(layout_path.name).suffix
        out_path = TMP / f"{stem}_완성_{ts}{suffix}"

        with st.spinner("색칠 진행 중..."):
            fl.run(
                src_path, layout_path, out_path,
                do_style=do_style,
                sheet_name=(sheet_name.strip() or None),
            )

        st.success(f"완료: {out_path.name}")
        mime = ("application/vnd.ms-excel.sheet.macroEnabled.12" if suffix == ".xlsm"
                else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button(
            "📥 결과 파일 다운로드",
            data=out_path.read_bytes(),
            file_name=out_path.name,
            mime=mime,
            width='stretch',
        )
    except Exception as e:
        st.error(f"오류: {e}")
        st.code(traceback.format_exc())

st.divider()
with st.expander("도움말"):
    st.markdown("""
- **원본 데이터**: 첫 열에 로케이션ID(`A11-01-10` 형태), 그다음 열에 품목코드·품목명이 있는 엑셀
- **레이아웃 파일**: 기본은 저장소에 내장된 파일 사용. 다른 배치도를 쓰려면 옵션에서 업로드
- **시각 디자인 적용 해제**: 값만 채우고 색은 안 넣음
- 시트 이름은 기본적으로 `재고조사`를 찾고, 없으면 첫 시트 사용
""")
