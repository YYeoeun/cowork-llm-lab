from abc import ABC, abstractmethod


class BaseModelClient(ABC):
    """모든 모델 클라이언트의 공통 인터페이스."""

    @abstractmethod
    def chat(self, messages: list[dict], system: str | None = None) -> str: ...
