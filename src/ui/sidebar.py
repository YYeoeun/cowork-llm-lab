import streamlit as st

MODEL_OPTIONS = [
    "openai/gpt-4o-mini",
    "openai/gpt-4o",
    "gemini/gemini-1.5-flash",
    "ollama/llama3.1",
    "ollama/qwen2.5",
    "ollama/mistral",
]


def render_sidebar() -> dict:
    """모델/서버 선택 사이드바를 렌더링하고 선택값을 반환."""
    with st.sidebar:
        st.header("설정")
        model = st.selectbox("모델", MODEL_OPTIONS)
        server = st.selectbox("실행 서버", ["로컬", "원격"])
    return {"model": model, "server": server}
