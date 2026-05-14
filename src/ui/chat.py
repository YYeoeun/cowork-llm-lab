import streamlit as st

from src.models import get_client


def render_chat(config: dict) -> None:
    """채팅 영역: 멀티턴 대화 + 선택된 provider 호출."""
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

    with st.chat_message("assistant"):
        try:
            client = get_client(config["model"])
            with st.spinner("응답 생성 중..."):
                reply = client.chat(st.session_state.messages)
        except NotImplementedError:
            reply = f"`{config['model']}` 제공자는 아직 연동되지 않았습니다."
        except Exception as e:
            reply = f"오류가 발생했습니다: `{type(e).__name__}: {e}`"
        st.markdown(reply)

    st.session_state.messages.append({"role": "assistant", "content": reply})
