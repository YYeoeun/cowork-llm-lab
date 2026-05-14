import os

from ollama import Client

from .base import BaseModelClient
from ..prompts.templates import SYSTEM_PROMPT


def list_models() -> list[str]:
    """Ollama 서버(OLLAMA_HOST or localhost)의 설치된 모델 목록."""
    client = Client(host=os.environ.get("OLLAMA_HOST"))
    response = client.list()
    # ollama-python 버전에 따라 dict 또는 Pydantic 모델 반환. 둘 다 지원.
    models_list = response.get("models", []) if isinstance(response, dict) else getattr(response, "models", [])
    names: list[str] = []
    for m in models_list:
        if isinstance(m, dict):
            name = m.get("model") or m.get("name")
        else:
            name = getattr(m, "model", None) or getattr(m, "name", None)
        if name:
            names.append(name)
    return sorted(names)


class OllamaClient(BaseModelClient):
    def __init__(self, model: str = "llama3.1") -> None:
        self.model = model
        self._client = Client(host=os.environ.get("OLLAMA_HOST"))

    def chat(self, messages: list[dict], system: str | None = None) -> str:
        payload = [{"role": "system", "content": system or SYSTEM_PROMPT}, *messages]
        response = self._client.chat(model=self.model, messages=payload)
        return response["message"]["content"]
