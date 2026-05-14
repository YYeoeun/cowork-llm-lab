from io import BytesIO

import pandas as pd


def load_excel(path) -> pd.DataFrame:
    return pd.read_excel(path)


def merge_excels(paths: list) -> pd.DataFrame:
    return pd.concat([load_excel(p) for p in paths], ignore_index=True)


def save_excel(df: pd.DataFrame, path) -> None:
    df.to_excel(path, index=False)


def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """DataFrame을 xlsx 바이트로 직렬화 (st.download_button용)."""
    buf = BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def df_to_markdown(df: pd.DataFrame) -> str:
    """tabulate 의존 없이 DataFrame을 GitHub-flavored markdown 표로 변환."""
    cols = [str(c) for c in df.columns]
    header = "| " + " | ".join(cols) + " |"
    sep = "| " + " | ".join("---" for _ in cols) + " |"
    rows = [
        "| " + " | ".join("" if pd.isna(v) else str(v) for v in row) + " |"
        for row in df.itertuples(index=False, name=None)
    ]
    return "\n".join([header, sep, *rows])
