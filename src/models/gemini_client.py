import os

import google.generativeai as genai

from .base import BaseModelClient
from ..prompts.templates import SYSTEM_PROMPT


class GeminiClient(BaseModelClient):
    def __init__(self, model: str = "gemini-1.5-flash") -> None:
        genai.configure(api_key=os.environ.get("GOOGLE_API_KEY"))
        self.model_name = model

    def chat(self, messages: list[dict]) -> str:
        model = genai.GenerativeModel(self.model_name, system_instruction=SYSTEM_PROMPT)
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
