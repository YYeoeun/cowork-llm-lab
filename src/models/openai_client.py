from .base import BaseModelClient


class OpenAIClient(BaseModelClient):
    def chat(self, messages: list[dict]) -> str:
        raise NotImplementedError
