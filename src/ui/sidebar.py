import streamlit as st


def render_sidebar() -> dict:
    """모델/서버 선택 사이드바를 렌더링하고 선택값을 반환."""
    with st.sidebar:
        st.header("설정")
        model = st.selectbox("모델", ["(placeholder)"])
        server = st.selectbox("실행 서버", ["로컬", "원격"])
    return {"model": model, "server": server}
