import os

from openai import OpenAI

from .base import BaseModelClient
from ..prompts.templates import SYSTEM_PROMPT

_EXCLUDE_SUBSTR = ("audio", "realtime", "tts", "transcribe", "search", "embed", "image", "moderation")


def list_models() -> list[str]:
    """OpenAI 계정에서 사용 가능한 chat 모델 ID 목록."""
    client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
    ids = [m.id for m in client.models.list().data]
    return sorted(
        m for m in ids
        if m.startswith(("gpt-", "o1", "o3", "chatgpt"))
        and not any(s in m for s in _EXCLUDE_SUBSTR)
    )


class OpenAIClient(BaseModelClient):
    def __init__(self, model: str = "gpt-4o-mini") -> None:
        self.model = model
        self._client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

    def chat(self, messages: list[dict], system: str | None = None) -> str:
        payload = [{"role": "system", "content": system or SYSTEM_PROMPT}, *messages]
        response = self._client.chat.completions.create(
            model=self.model,
            messages=payload,
        )
        return response.choices[0].message.content or ""
