"""엑셀 처리 파이프라인: LLM이 생성한 pandas 코드를 실행해 결과를 만든다."""

import re
from pathlib import Path
from typing import Any

import pandas as pd

from src.files.excel import load_excel
from src.models.base import BaseModelClient
from src.prompts.templates import CODE_SYSTEM_PROMPT

_CODE_RE = re.compile(r"```(?:python)?\s*\n(.*?)```", re.DOTALL)


def build_file_context(paths: list[Path]) -> str:
    """업로드된 각 엑셀 파일의 shape/columns/dtypes/미리보기를 텍스트로 요약."""
    lines: list[str] = []
    for i, p in enumerate(paths):
        df = load_excel(p)
        lines.append(f"### dfs[{i}] = {p.name}  (shape={df.shape})")
        lines.append(f"columns: {list(df.columns)}")
        lines.append(f"dtypes: {{ {', '.join(f'{c}: {t}' for c, t in df.dtypes.astype(str).items())} }}")
        lines.append("preview:")
        lines.append(df.head(3).to_csv(index=False))
    return "\n".join(lines)


def extract_code(text: str) -> str:
    """모델 응답에서 ```python 블록을 추출. 없으면 본문 전체 반환."""
    m = _CODE_RE.search(text)
    return m.group(1) if m else text


def run_code(code: str, paths: list[Path]) -> Any:
    """`dfs` 리스트와 `pd`를 노출한 네임스페이스에서 코드 실행 후 `result` 반환."""
    dfs = [load_excel(p) for p in paths]
    ns: dict = {"pd": pd, "dfs": dfs}
    exec(code, ns)
    return ns.get("result")


def _attempt(messages: list[dict], paths: list[Path], client: BaseModelClient) -> dict:
    """LLM 호출 1회 + run_code 1회. 에러는 raise하지 않고 dict로 반환."""
    response = client.chat(messages, system=CODE_SYSTEM_PROMPT)
    code = extract_code(response)
    try:
        result = run_code(code, paths)
        return {"response": response, "code": code, "result": result, "error": None}
    except Exception as e:
        return {
            "response": response,
            "code": code,
            "result": None,
            "error": f"{type(e).__name__}: {e}",
        }


def process(prompt: str, paths: list[Path], client: BaseModelClient) -> dict:
    """엑셀 파이프라인 1회 실행 (+ 실패 시 1회 자동 재시도).

    반환 dict 키:
      - response: 모델 최종 응답 (재시도된 경우 두 번째 응답)
      - code: 최종 실행 대상 코드
      - result: 실행 결과 (최종 실패 시 None)
      - error: 최종 에러 문자열 (성공 시 None)
      - retried: 재시도 발생 여부
      - first_error: 첫 시도 에러 (재시도된 경우만, 아니면 None)
    """
    context = build_file_context(paths)
    messages = [{"role": "user", "content": f"{context}\n\n사용자 요청:\n{prompt}"}]

    first = _attempt(messages, paths, client)
    if first["error"] is None:
        return {**first, "retried": False, "first_error": None}

    # 1회 재시도: 에러를 LLM에 피드백하여 수정된 코드 요청
    messages.append({"role": "assistant", "content": first["response"]})
    messages.append(
        {
            "role": "user",
            "content": (
                f"위 코드를 실행했더니 다음 에러가 발생했습니다:\n```\n{first['error']}\n```\n"
                f"수정된 코드만 다시 작성해주세요."
            ),
        }
    )
    second = _attempt(messages, paths, client)
    return {**second, "retried": True, "first_error": first["error"]}
