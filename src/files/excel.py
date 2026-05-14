import pandas as pd


def load_excel(path) -> pd.DataFrame:
    raise NotImplementedError


def merge_excels(paths: list) -> pd.DataFrame:
    raise NotImplementedError


def save_excel(df: pd.DataFrame, path) -> None:
    raise NotImplementedError
