from collections import defaultdict
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="cowork-llm-lab", layout="wide")

st.title("cowork-llm-lab")
st.caption("LLM과 함께 협업하는 Streamlit 작업 도구")


def parse_numeric(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).replace(",", "").strip()
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def is_numeric_series(series):
    if pd.api.types.is_numeric_dtype(series):
        return True
    non_null = series.dropna()
    if non_null.empty:
        return False
    return all(parse_numeric(v) is not None for v in non_null)


def stringify_objects(df):
    out = df.copy()
    for col in out.columns:
        if out[col].dtype == "object":
            out[col] = out[col].astype(str)
    return out


def detect_header_row(raw, max_scan=15):
    if raw.empty:
        return 0
    limit = min(max_scan, len(raw))
    densities = [raw.iloc[i].notna().sum() for i in range(limit)]
    peak = max(densities)
    if peak == 0:
        return 0
    threshold = max(2, int(peak * 0.7))

    best_idx = 0
    best_score = -1
    for i in range(limit):
        row = raw.iloc[i].dropna()
        if len(row) < threshold:
            continue
        if any(parse_numeric(v) is not None for v in row):
            continue
        if i + 1 >= len(raw):
            continue
        next_row = raw.iloc[i + 1].dropna()
        if len(next_row) == 0:
            continue
        next_numeric = sum(1 for v in next_row if parse_numeric(v) is not None)
        score = len(row) + next_numeric * 2
        if score > best_score:
            best_score = score
            best_idx = i
    return best_idx if best_score >= 0 else 0


def read_sheet(excel, sheet_name):
    raw = excel.parse(sheet_name, header=None)
    if raw.empty:
        return raw, 0
    header_idx = detect_header_row(raw)
    header_vals = raw.iloc[header_idx]
    valid = header_vals.notna()
    if not valid.any():
        df = raw.iloc[header_idx + 1 :].copy().reset_index(drop=True)
        return df, header_idx
    df = raw.iloc[header_idx + 1 :, valid.values].copy()
    df.columns = header_vals[valid].astype(str).values
    df = df.dropna(how="all").reset_index(drop=True)
    return df, header_idx


def sanitize_sheet_name(name, used):
    for ch in "[]:*?/\\'":
        name = name.replace(ch, "_")
    name = name.strip()[:31] or "sheet"
    base = name
    i = 2
    while name in used:
        suffix = f"_{i}"
        name = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(name)
    return name


def to_cell_value(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, float) and val.is_integer():
        return int(val)
    return val


def write_dataframe(ws, df):
    for col_idx, col_name in enumerate(df.columns, 1):
        ws.cell(row=1, column=col_idx, value=str(col_name))
    for row_idx, row in enumerate(df.itertuples(index=False, name=None), 2):
        for col_idx, val in enumerate(row, 1):
            cell_val = to_cell_value(val)
            if cell_val is not None:
                ws.cell(row=row_idx, column=col_idx, value=cell_val)


def classify_columns(signature, members):
    numeric = [
        col
        for col in signature
        if all(is_numeric_series(df[col]) for _, _, df in members)
    ]
    keys = [col for col in signature if col not in numeric]
    return keys, numeric


def union_keys(key_cols, members):
    if not key_cols:
        return None
    parts = [df[key_cols] for _, _, df in members]
    union = pd.concat(parts, ignore_index=True).drop_duplicates()
    return union.sort_values(key_cols, kind="stable").reset_index(drop=True)


def realign(df, signature, key_cols, numeric_cols, all_keys, max_len):
    sub = df[list(signature)].copy()
    for col in numeric_cols:
        sub[col] = sub[col].apply(parse_numeric)
    if key_cols:
        sub = sub.drop_duplicates(subset=key_cols, keep="first")
        aligned = all_keys.merge(sub, on=list(key_cols), how="left")
        aligned = aligned[list(signature)]
    else:
        aligned = sub.reindex(range(max_len))
    return aligned


def build_merged_workbook(groups):
    wb = Workbook()
    wb.remove(wb.active)
    used = set()

    for group_idx, (signature, members) in enumerate(groups.items(), 1):
        if len(members) == 1:
            file_name, sheet_name, df = members[0]
            stem = file_name.rsplit(".", 1)[0]
            ws_name = sanitize_sheet_name(f"{stem}_{sheet_name}", used)
            ws = wb.create_sheet(ws_name)
            write_dataframe(ws, df)
            continue

        key_cols, numeric_cols = classify_columns(signature, members)
        all_keys = union_keys(key_cols, members)
        max_len = max(len(df) for _, _, df in members)
        n_rows = len(all_keys) if key_cols else max_len

        realigned = [
            (f, s, realign(df, signature, key_cols, numeric_cols, all_keys, max_len))
            for f, s, df in members
        ]
        source_names = []
        for file_name, sheet_name, _ in realigned:
            stem = file_name.rsplit(".", 1)[0]
            source_names.append(sanitize_sheet_name(f"{stem}_{sheet_name}", used))

        summary_name = sanitize_sheet_name(f"summary_{group_idx}", used)
        summary_ws = wb.create_sheet(summary_name)
        for col_idx, col_name in enumerate(signature, 1):
            summary_ws.cell(row=1, column=col_idx, value=str(col_name))

        for row_idx in range(n_rows):
            excel_row = row_idx + 2
            for col_idx, col_name in enumerate(signature, 1):
                col_letter = get_column_letter(col_idx)
                if col_name in numeric_cols:
                    refs = [f"'{sn}'!{col_letter}{excel_row}" for sn in source_names]
                    summary_ws.cell(
                        row=excel_row,
                        column=col_idx,
                        value=f"=SUM({','.join(refs)})",
                    )
                elif key_cols:
                    val = all_keys.iloc[row_idx][col_name]
                    summary_ws.cell(
                        row=excel_row, column=col_idx, value=to_cell_value(val)
                    )

        for (file_name, sheet_name, aligned), ws_name in zip(realigned, source_names):
            ws = wb.create_sheet(ws_name)
            write_dataframe(ws, aligned)

    return wb


with st.sidebar:
    st.header("엑셀 업로드")
    uploaded_files = st.file_uploader(
        "통합할 엑셀 파일을 선택하세요",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )

if not uploaded_files:
    st.info("왼쪽 사이드바에서 엑셀 파일을 업로드하세요.")
    st.stop()

all_sheets = []
header_indices = {}
for file in uploaded_files:
    try:
        excel = pd.ExcelFile(file)
    except Exception as e:
        st.error(f"{file.name}: 파일을 읽지 못했습니다 ({e})")
        continue
    for sheet_name in excel.sheet_names:
        df, header_idx = read_sheet(excel, sheet_name)
        all_sheets.append((file.name, sheet_name, df))
        header_indices[(file.name, sheet_name)] = header_idx

if not all_sheets:
    st.stop()

with st.expander("원본 미리보기", expanded=False):
    file_names = list(dict.fromkeys(f for f, _, _ in all_sheets))
    for file_name in file_names:
        st.markdown(f"**{file_name}**")
        members = [(s, df) for f, s, df in all_sheets if f == file_name]
        tabs = st.tabs([s for s, _ in members])
        for tab, (sheet_name, df) in zip(tabs, members):
            with tab:
                h_idx = header_indices.get((file_name, sheet_name), 0)
                st.caption(
                    f"{len(df)} rows × {len(df.columns)} cols · 감지된 헤더 행: {h_idx} (0-based)"
                )
                st.dataframe(stringify_objects(df), width="stretch")

groups = defaultdict(list)
for file_name, sheet_name, df in all_sheets:
    groups[tuple(df.columns)].append((file_name, sheet_name, df))

st.header("매칭 그룹")
st.caption(
    "컬럼 구조가 동일한 시트끼리 묶어, 텍스트 컬럼은 키로 두고 숫자 컬럼만 통합 대상으로 분류합니다."
)
for group_idx, (signature, members) in enumerate(groups.items(), 1):
    sources = [f"{f} / {s}" for f, s, _ in members]
    st.subheader(f"그룹 {group_idx} · {len(signature)}컬럼 · {len(members)}시트")
    st.caption(" · ".join(sources))
    key_cols, numeric_cols = classify_columns(signature, members)
    st.write(
        f"키 컬럼: {key_cols or '(없음)'}  ·  숫자 컬럼: {numeric_cols or '(없음)'}"
    )

st.divider()
st.header("통합 엑셀")
st.caption(
    "같은 구조 시트끼리 묶어 그룹별로 합산 시트(SUM 수식)를 앞에 두고, "
    "원본 시트는 키 기준으로 정렬해 그 뒤에 배치합니다. 단일 시트 그룹은 원본만 복사합니다."
)

if st.button("통합 엑셀 생성", type="primary"):
    wb = build_merged_workbook(groups)
    buffer = BytesIO()
    wb.save(buffer)
    st.session_state["merged_xlsx"] = buffer.getvalue()
    st.success("생성 완료")

if "merged_xlsx" in st.session_state:
    st.download_button(
        "merged.xlsx 다운로드",
        data=st.session_state["merged_xlsx"],
        file_name="merged.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
