class RemoteRunner:
    """RTX 5090 / Spark 등 원격 GPU 서버에서 모델 추론 실행을 담당."""

    def run(self, model: str, payload: dict) -> dict:
        raise NotImplementedError
