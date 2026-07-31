# -*- coding: utf-8 -*-
"""각 페이지 상단에 넣는 공용 '사용 설명서' 헬퍼."""
import streamlit as st


def show_help(sections: dict, title: str = "📖 사용 설명서", expanded: bool = False):
    """페이지 최상단 접이식 설명서.
    sections: {"목적": "...", "필요한 파일": "...", "사용 순서": "1. ...", "결과": "..."} 형태.
    """
    with st.expander(title, expanded=expanded):
        for k, v in sections.items():
            st.markdown(f"**{k}**")
            st.markdown(v)
            st.markdown("")
