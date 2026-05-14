from .base import BaseModelClient

# provider별 fallback (동적 조회 실패 시 사용). Ollama는 설치 모델이 없으면 빈 게 정확.
_FALLBACK = {
    "openai": ["gpt-4o-mini", "gpt-4o"],
    "gemini": ["gemini-1.5-flash"],
    "ollama": [],
}


def list_all_models() -> list[str]:
    """각 provider에서 사용 가능한 모델을 'provider/name' 형식으로 반환.

    provider 조회 실패 시 해당 provider의 fallback 리스트 사용.
    전부 실패하면 안전한 기본 옵션을 반환해 UI가 빈 상태가 되지 않게 함.
    """
    result: list[str] = []
    for provider, module_name in (
        ("openai", "openai_client"),
        ("gemini", "gemini_client"),
        ("ollama", "ollama_client"),
    ):
        try:
            mod = __import__(f"src.models.{module_name}", fromlist=["list_models"])
            names = mod.list_models()
        except Exception:
            names = _FALLBACK[provider]
        result.extend(f"{provider}/{n}" for n in names)
    if not result:
        return ["openai/gpt-4o-mini"]
    return result


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
