import pandas as pd
import streamlit as st

from src.files.excel import df_to_markdown, df_to_xlsx_bytes
from src.models import get_client
from src.prompts.pipeline import process


def render_chat(config: dict) -> None:
    """채팅 영역: 일반 채팅 + (파일 있을 때) 엑셀 처리 파이프라인."""
    st.title("cowork-llm-lab")

    files = config.get("files", [])
    if files:
        st.caption(f"📂 사용 가능한 파일 ({len(files)}): {', '.join(f.name for f in files)}")
        st.warning(
            "⚠️ 파일이 업로드된 상태에서는 모델이 생성한 Python 코드가 로컬에서 실행됩니다."
        )

    if "messages" not in st.session_state:
        st.session_state.messages = []
    if "artifacts" not in st.session_state:
        st.session_state.artifacts = {}

    for i, msg in enumerate(st.session_state.messages):
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
            if i in st.session_state.artifacts:
                _render_artifact(st.session_state.artifacts[i], key=f"art_{i}")

    prompt = st.chat_input("프롬프트를 입력하세요")
    if not prompt:
        return

    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        artifact: dict | None = None
        try:
            client = get_client(config["model"])
            with st.spinner("응답 생성 중..."):
                if files:
                    artifact = process(prompt, files, client)
                    reply = artifact["response"]
                else:
                    reply = client.chat(st.session_state.messages)
        except NotImplementedError:
            reply = f"`{config['model']}` 제공자는 아직 연동되지 않았습니다."
        except Exception as e:
            reply = f"오류가 발생했습니다: `{type(e).__name__}: {e}`"

        st.markdown(reply)
        st.session_state.messages.append({"role": "assistant", "content": reply})
        if artifact is not None:
            idx = len(st.session_state.messages) - 1
            st.session_state.artifacts[idx] = artifact
            _render_artifact(artifact, key=f"art_{idx}")


def _render_artifact(artifact: dict, key: str) -> None:
    """파이프라인 결과(result) + xlsx/md 다운로드 버튼을 렌더링."""
    if artifact.get("retried"):
        st.caption(f"🔄 1회 재시도됨 (첫 시도 에러: `{artifact.get('first_error')}`)")

    error = artifact.get("error")
    if error:
        st.error(f"코드 실행 실패: `{error}`")
        with st.expander("실행을 시도한 코드 보기", expanded=False):
            st.code(artifact.get("code", "(코드 없음)"), language="python")
        st.caption(
            "💡 동일 프롬프트로 한 번 더 보내거나, 프롬프트에 컬럼명/원하는 출력 형태를 더 구체적으로 적어보세요."
        )
        return

    result = artifact.get("result")

    if isinstance(result, pd.DataFrame):
        st.dataframe(result, use_container_width=True)
        col_xlsx, col_md = st.columns(2)
        col_xlsx.download_button(
            "xlsx 다운로드",
            data=df_to_xlsx_bytes(result),
            file_name="result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key}_xlsx",
        )
        col_md.download_button(
            "markdown 다운로드",
            data=df_to_markdown(result),
            file_name="result.md",
            mime="text/markdown",
            key=f"{key}_md",
        )
    elif result is not None:
        text = str(result)
        st.code(text)
        st.download_button(
            "markdown 다운로드",
            data=text,
            file_name="result.md",
            mime="text/markdown",
            key=f"{key}_md",
        )
    else:
        st.info("코드가 실행되었지만 `result` 변수가 정의되지 않았습니다.")
