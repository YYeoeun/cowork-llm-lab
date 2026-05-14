import streamlit as st

from src.files.storage import save_upload, list_uploads, delete_upload
from src.models import list_all_models


@st.cache_data(ttl=3600, show_spinner="모델 목록 조회 중...")
def _fetch_models() -> list[str]:
    return list_all_models()


def render_sidebar() -> dict:
    """모델/서버 선택 + 엑셀 파일 관리 사이드바."""
    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0
    if "saved_uploads" not in st.session_state:
        st.session_state.saved_uploads = set()

    with st.sidebar:
        st.header("설정")
        options = _fetch_models()
        model = st.selectbox("모델", options)
        if st.button("🔄 모델 목록 새로고침", use_container_width=True):
            _fetch_models.clear()
            st.rerun()
        server = st.selectbox("실행 서버", ["로컬", "원격"])

        st.divider()
        st.subheader("엑셀 파일")
        uploaded = st.file_uploader(
            "업로드",
            type=["xlsx"],
            accept_multiple_files=True,
            key=f"uploader_{st.session_state.uploader_key}",
            label_visibility="collapsed",
        )
        if uploaded:
            for f in uploaded:
                if f.name not in st.session_state.saved_uploads:
                    save_upload(f)
                    st.session_state.saved_uploads.add(f.name)

        files = list_uploads()
        if not files:
            st.caption("업로드된 파일 없음")
        else:
            for f in files:
                col1, col2 = st.columns([4, 1])
                col1.text(f.name)
                if col2.button("삭제", key=f"del_{f.name}"):
                    delete_upload(f.name)
                    st.session_state.saved_uploads.discard(f.name)
                    st.session_state.uploader_key += 1
                    st.rerun()

    return {"model": model, "server": server, "files": files}
