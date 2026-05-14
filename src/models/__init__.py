from .base import BaseModelClient


def get_client(model_id: str) -> BaseModelClient:
    """`provider/model` 식별자(예: 'openai/gpt-4o-mini')를 받아 클라이언트를 반환.

    각 provider 모듈은 lazy import — 사용 안 하는 provider의 SDK는 설치 불필요.
    """
    provider, _, name = model_id.partition("/")
    if provider == "openai":
        from .openai_client import OpenAIClient
        return OpenAIClient(model=name)
    if provider == "gemini":
        from .gemini_client import GeminiClient
        return GeminiClient(model=name)
    if provider == "ollama":
        from .ollama_client import OllamaClient
        return OllamaClient(model=name)
    raise ValueError(f"Unknown provider in model_id: {model_id!r}")
