from .base import BaseModelClient


class GeminiClient(BaseModelClient):
    def chat(self, messages: list[dict]) -> str:
        raise NotImplementedError
