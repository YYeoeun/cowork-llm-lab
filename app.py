import streamlit as st
from dotenv import load_dotenv

from src.ui.sidebar import render_sidebar
from src.ui.chat import render_chat


def main() -> None:
    load_dotenv()
    st.set_page_config(page_title="cowork-llm-lab", layout="wide")
    config = render_sidebar()
    render_chat(config)


if __name__ == "__main__":
    main()
