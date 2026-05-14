"""엑셀 처리 파이프라인: LLM이 생성한 pandas 코드를 실행해 결과를 만든다."""

import re
from pathlib import Path
from typing import Any

import pandas as pd

from src.files.excel import load_excel
from src.models.base import BaseModelClient
from src.prompts.templates import CODE_SYSTEM_PROMPT

_CODE_RE = re.compile(r"```(?:python)?\s*\n(.*?)```", re.DOTALL)


def _build_combined(paths: list[Path], dfs: list[pd.DataFrame]) -> pd.DataFrame:
    """모든 dfs를 concat. 각 행에 `_source_file` 컬럼으로 출처 표시."""
    parts = [df.assign(_source_file=p.name) for p, df in zip(paths, dfs)]
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


def _split_columns(combined: pd.DataFrame) -> tuple[list[str], list[str]]:
    """combined의 컬럼을 (text_cols, numeric_cols)로 분리. `_source_file`은 제외."""
    cols = [c for c in combined.columns if c != "_source_file"]
    text_cols = [c for c in cols if not pd.api.types.is_numeric_dtype(combined[c])]
    numeric_cols = [c for c in cols if pd.api.types.is_numeric_dtype(combined[c])]
    return text_cols, numeric_cols


def build_file_context(paths: list[Path], dfs: list[pd.DataFrame]) -> str:
    """업로드된 각 엑셀 파일의 schema/preview + 교차 파일 컬럼 겹침 요약."""
    lines: list[str] = []

    # 파일별 요약
    for i, (p, df) in enumerate(zip(paths, dfs)):
        lines.append(f"### dfs[{i}] = {p.name}  (shape={df.shape})")
        lines.append(f"columns: {list(df.columns)}")
        lines.append(f"dtypes: {{ {', '.join(f'{c}: {t}' for c, t in df.dtypes.astype(str).items())} }}")
        lines.append("preview:")
        lines.append(df.head(3).to_csv(index=False))

    # 교차 파일 컬럼 겹침 분석 (파일이 2개 이상일 때만)
    if len(dfs) > 1:
        col_sets = [set(df.columns) for df in dfs]
        common = set.intersection(*col_sets)
        union = set.union(*col_sets)
        partial = sorted(
            (col, [p.name for p, s in zip(paths, col_sets) if col in s])
            for col in union
            if 1 < sum(1 for s in col_sets if col in s) < len(dfs)
        )
        lines.append("### 컬럼 겹침 분석")
        lines.append(
            f"모든 파일 공통 컬럼: {sorted(common) if common else '없음'}"
        )
        if partial:
            lines.append("일부 파일에만 있는 공유 컬럼:")
            for col, files in partial:
                lines.append(f"  - {col}: {files}")

    # 사전 빌드된 변수 안내
    combined = _build_combined(paths, dfs)
    text_cols, numeric_cols = _split_columns(combined)
    lines.append("### 사전 빌드된 변수 (네임스페이스에 이미 존재)")
    lines.append("- `combined`: 모든 dfs를 concat한 DataFrame (`_source_file` 컬럼으로 출처 표시)")
    lines.append(f"  - `text_cols` (내용/비-숫자): {text_cols}")
    lines.append(f"  - `numeric_cols` (숫자): {numeric_cols}")
    lines.append("- `rapidfuzz`: 퍼지 문자열 매칭 라이브러리 (예: rapidfuzz.fuzz.ratio, rapidfuzz.process.extractOne)")

    return "\n".join(lines)


def extract_code(text: str) -> str:
    """모델 응답에서 ```python 블록을 추출. 없으면 본문 전체 반환."""
    m = _CODE_RE.search(text)
    return m.group(1) if m else text


def run_code(code: str, paths: list[Path], dfs: list[pd.DataFrame]) -> Any:
    """`dfs`, `combined`, `text_cols`, `numeric_cols`, `pd`, `rapidfuzz`(설치 시) 네임스페이스."""
    combined = _build_combined(paths, dfs)
    text_cols, numeric_cols = _split_columns(combined)

    ns: dict = {
        "pd": pd,
        "dfs": dfs,
        "combined": combined,
        "text_cols": text_cols,
        "numeric_cols": numeric_cols,
    }

    # rapidfuzz는 lazy import — 미설치 환경에서도 일반 코드는 동작
    try:
        import rapidfuzz
        ns["rapidfuzz"] = rapidfuzz
    except ImportError:
        pass

    exec(code, ns)
    return ns.get("result")


def _attempt(
    messages: list[dict],
    paths: list[Path],
    dfs: list[pd.DataFrame],
    client: BaseModelClient,
) -> dict:
    """LLM 호출 1회 + run_code 1회. 에러는 raise하지 않고 dict로 반환."""
    response = client.chat(messages, system=CODE_SYSTEM_PROMPT)
    code = extract_code(response)
    try:
        result = run_code(code, paths, dfs)
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
    dfs = [load_excel(p) for p in paths]
    context = build_file_context(paths, dfs)
    messages = [{"role": "user", "content": f"{context}\n\n사용자 요청:\n{prompt}"}]

    first = _attempt(messages, paths, dfs, client)
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
    second = _attempt(messages, paths, dfs, client)
    return {**second, "retried": True, "first_error": first["error"]}
