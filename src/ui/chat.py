import streamlit as st


def render_chat(config: dict) -> None:
    """채팅 영역(업로드 + 대화) placeholder."""
    st.title("cowork-llm-lab")
    st.file_uploader("엑셀 업로드", type=["xlsx"], accept_multiple_files=True)
    st.chat_input("프롬프트를 입력하세요")
