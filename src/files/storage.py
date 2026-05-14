from pathlib import Path

DATA_DIR = Path("data")


def save_upload(file) -> Path:
    """Streamlit UploadedFile을 data/ 디렉토리에 저장하고 경로 반환."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    target = DATA_DIR / file.name
    target.write_bytes(file.getbuffer())
    return target


def list_uploads() -> list[Path]:
    """data/ 디렉토리의 업로드 파일 목록을 정렬해 반환 (.gitkeep 제외)."""
    if not DATA_DIR.exists():
        return []
    return sorted(
        p for p in DATA_DIR.iterdir() if p.is_file() and p.name != ".gitkeep"
    )


def delete_upload(name: str) -> None:
    """파일명으로 업로드 파일 삭제. 없으면 무시."""
    path = DATA_DIR / name
    if path.is_file():
        path.unlink()
