import os

import google.generativeai as genai

from .base import BaseModelClient
from ..prompts.templates import SYSTEM_PROMPT


class GeminiClient(BaseModelClient):
    def __init__(self, model: str = "gemini-1.5-flash") -> None:
        genai.configure(api_key=os.environ.get("GOOGLE_API_KEY"))
        self.model_name = model

    def chat(self, messages: list[dict], system: str | None = None) -> str:
        # Gemini의 system_instruction은 모델 생성 시 고정이라 호출마다 인스턴스화.
        model = genai.GenerativeModel(
            self.model_name, system_instruction=system or SYSTEM_PROMPT
        )
        history = [
            {
                "role": "model" if m["role"] == "assistant" else "user",
                "parts": [m["content"]],
            }
            for m in messages[:-1]
        ]
        session = model.start_chat(history=history)
        response = session.send_message(messages[-1]["content"])
        return response.text
