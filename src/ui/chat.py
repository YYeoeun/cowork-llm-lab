import streamlit as st


def render_chat(config: dict) -> None:
    """채팅 영역: 파일 업로드 + 멀티턴 대화 히스토리."""
    st.title("cowork-llm-lab")
    st.file_uploader("엑셀 업로드", type=["xlsx"], accept_multiple_files=True)

    if "messages" not in st.session_state:
        st.session_state.messages = []

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    prompt = st.chat_input("프롬프트를 입력하세요")
    if not prompt:
        return

    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    reply = _generate_reply(prompt, config)
    st.session_state.messages.append({"role": "assistant", "content": reply})
    with st.chat_message("assistant"):
        st.markdown(reply)


def _generate_reply(prompt: str, config: dict) -> str:
    """모델 응답 생성. 모델 연동 전까지는 placeholder 반환."""
    model = config.get("model", "(unknown)")
    server = config.get("server", "(unknown)")
    return (
        f"[준비 중] `{model}` / `{server}` 연동이 완료되면 응답합니다.\n\n"
        f"입력: {prompt}"
    )
