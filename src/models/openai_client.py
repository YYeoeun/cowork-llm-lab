import os

from openai import OpenAI

from .base import BaseModelClient
from ..prompts.templates import SYSTEM_PROMPT


class OpenAIClient(BaseModelClient):
    def __init__(self, model: str = "gpt-4o-mini") -> None:
        self.model = model
        self._client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

    def chat(self, messages: list[dict]) -> str:
        payload = [{"role": "system", "content": SYSTEM_PROMPT}, *messages]
        response = self._client.chat.completions.create(
            model=self.model,
            messages=payload,
        )
        return response.choices[0].message.content or ""
