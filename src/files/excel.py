import pandas as pd


def load_excel(path) -> pd.DataFrame:
    return pd.read_excel(path)


def merge_excels(paths: list) -> pd.DataFrame:
    return pd.concat([load_excel(p) for p in paths], ignore_index=True)


def save_excel(df: pd.DataFrame, path) -> None:
    df.to_excel(path, index=False)
