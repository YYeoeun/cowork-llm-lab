from pathlib import Path

DATA_DIR = Path("data")


def save_upload(file) -> Path:
    raise NotImplementedError


def list_uploads() -> list[Path]:
    raise NotImplementedError


def delete_upload(name: str) -> None:
    raise NotImplementedError
