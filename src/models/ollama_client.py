import os

from ollama import Client

from .base import BaseModelClient
from ..prompts.templates import SYSTEM_PROMPT


class OllamaClient(BaseModelClient):
    def __init__(self, model: str = "llama3.1") -> None:
        self.model = model
        self._client = Client(host=os.environ.get("OLLAMA_HOST"))

    def chat(self, messages: list[dict], system: str | None = None) -> str:
        payload = [{"role": "system", "content": system or SYSTEM_PROMPT}, *messages]
        response = self._client.chat(model=self.model, messages=payload)
        return response["message"]["content"]
