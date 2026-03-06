from __future__ import annotations

import re
from datetime import timedelta
from typing import Tuple

import pandas as pd

from .helpers import AppError, ensure_columns

TICKET_REQUIRED_COLUMNS = [
    "创建时间",
    "关联提出人",
    "工单分类",
    "工单标签",
    "计数",
    "预计工时",
    "客服结论",
    "客服补充",
]

QA_REQUIRED_COLUMNS = [
    "客服",
    "等级",
    "所属周次",
    "质检会话占比",
]

QA_VERSION_COLUMNS = [
    "版本",
    "版本（老版/新版）",
    "版本(老版/新版)",
]


def _remove_qa_example_rows(qa_df: pd.DataFrame) -> pd.DataFrame:
    if qa_df.empty:
        return qa_df

    first_col = qa_df.columns[0]
    first_col_text = qa_df[first_col].fillna("").astype(str).str.strip().str.lower()
    is_example = first_col_text.isin(["示例", "example"])

    if "客服" in qa_df.columns:
        agent_text = qa_df["客服"].fillna("").astype(str).str.strip().str.lower()
        is_example = is_example | agent_text.isin(["示例", "xxx"])

    return qa_df.loc[~is_example].copy()


def _read_excel(uploaded_file) -> pd.DataFrame:
    try:
        return pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as exc:  # noqa: BLE001
        raise AppError(f"Excel 解析失败：{exc}") from exc


def _match_week_range_col(col_name: str) -> bool:
    text = str(col_name).strip()
    pattern = r"^\d{4}[/-]\d{1,2}[/-]\d{1,2}\s*[-~至]\s*\d{4}[/-]\d{1,2}[/-]\d{1,2}$"
    return re.match(pattern, text) is not None


def _week_col_sort_key(col_name: str):
    text = str(col_name).strip()
    parts = re.split(r"[-~至]", text)
    if not parts:
        return pd.Timestamp.max
    start = pd.to_datetime(parts[0].strip(), errors="coerce")
    if pd.isna(start):
        return pd.Timestamp.max
    return start


def _extract_week_range(raw_col_name: str):
    text = str(raw_col_name).strip()
    # Remove pandas duplicate suffix like ".1", ".2".
    text = re.sub(r"\.\d+$", "", text)
    parts = re.split(r"[-~至]", text)
    if len(parts) != 2:
        return None
    start = pd.to_datetime(parts[0].strip(), errors="coerce")
    end = pd.to_datetime(parts[1].strip(), errors="coerce")
    if pd.isna(start) or pd.isna(end):
        return None
    return start.normalize(), end.normalize()


def _normalize_qa_wide_to_long(qa_df: pd.DataFrame) -> tuple[pd.DataFrame, list[tuple[int, str, pd.Timestamp, pd.Timestamp]]]:
    name_col = None
    for candidate in ["客服（人名）", "客服(人名)", "客服"]:
        if candidate in qa_df.columns:
            name_col = candidate
            break

    version_col = None
    for candidate in QA_VERSION_COLUMNS:
        if candidate in qa_df.columns:
            version_col = candidate
            break

    if name_col is None or "质检会话占比" not in qa_df.columns:
        raise AppError(
            "QA 结果格式不支持。请使用以下任一格式："
            "长表（客服, 等级, 所属周次, 质检会话占比）"
            "或宽表（客服（人名）+ 版本（老版/新版）+ 每周区间列 + 质检会话占比）。"
        )

    all_cols = list(qa_df.columns)

    # Accept repeated pair pattern:
    # [week_grade_col, ratio_col, week_grade_col.1, ratio_col.1, ...]
    # and also older format with one shared ratio column.
    grade_ratio_pairs = []
    shared_ratio_col = "质检会话占比" if "质检会话占比" in qa_df.columns else None
    ratio_like = [c for c in all_cols if str(c).startswith("质检会话占比")]

    for i, col in enumerate(all_cols):
        col_text = str(col).strip()
        if col == name_col:
            continue
        if version_col is not None and col == version_col:
            continue
        if col_text.startswith("质检会话占比"):
            continue
        if not _match_week_range_col(col_text):
            continue

        # Treat this as one week grade column.
        ratio_col = None
        if i + 1 < len(all_cols) and str(all_cols[i + 1]).startswith("质检会话占比"):
            ratio_col = all_cols[i + 1]
        elif shared_ratio_col is not None:
            ratio_col = shared_ratio_col
        elif ratio_like:
            ratio_col = ratio_like[0]

        grade_ratio_pairs.append((col, ratio_col))

    # De-duplicate while preserving order.
    uniq_pairs = []
    seen = set()
    for g_col, r_col in grade_ratio_pairs:
        if g_col in seen:
            continue
        seen.add(g_col)
        uniq_pairs.append((g_col, r_col))
    grade_ratio_pairs = uniq_pairs

    if not grade_ratio_pairs:
        raise AppError("QA 宽表中未找到周次等级列。")

    week_meta = []
    for idx, (grade_col, _) in enumerate(grade_ratio_pairs, start=1):
        parsed = _extract_week_range(str(grade_col))
        if parsed is None:
            raise AppError(
                f"QA 周次列格式无效：'{grade_col}'。"
                "应为日期区间，例如：'2026/02/23-2026/03/01'。"
            )
        week_meta.append((idx, str(grade_col), parsed[0], parsed[1]))

    rows = []
    for _, row in qa_df.iterrows():
        agent_name = str(row.get(name_col, "")).strip()
        if not agent_name or agent_name.lower() == "nan":
            continue
        qa_version = str(row.get(version_col, "")).strip() if version_col else ""

        for idx, (grade_col, ratio_col) in enumerate(grade_ratio_pairs, start=1):
            grade = row.get(grade_col)
            if pd.isna(grade) or str(grade).strip() == "":
                continue
            ratio = row.get(ratio_col) if ratio_col is not None else 0.0
            row_item = {
                "客服": agent_name,
                "等级": str(grade).strip(),
                "所属周次": f"W{idx}",
                "质检会话占比": ratio,
            }
            if qa_version:
                row_item["版本"] = qa_version
            rows.append(row_item)

    long_df = pd.DataFrame(rows)
    if long_df.empty:
        raise AppError("QA 文件格式转换后没有可用数据行。")
    return long_df, week_meta


def _normalize_qa_df(qa_df: pd.DataFrame) -> tuple[pd.DataFrame, list[tuple[int, str, pd.Timestamp, pd.Timestamp]]]:
    # Long format: keep as-is.
    if all(col in qa_df.columns for col in QA_REQUIRED_COLUMNS):
        return qa_df.copy(), []
    # Wide format: convert to long.
    return _normalize_qa_wide_to_long(qa_df)


def _validate_qa_week_ranges(ticket_df: pd.DataFrame, week_meta: list[tuple[int, str, pd.Timestamp, pd.Timestamp]]) -> None:
    if not week_meta:
        return

    created = pd.to_datetime(ticket_df["创建时间"], errors="coerce")
    if created.isna().any():
        raise AppError("字段「创建时间」包含无效值，无法校验 QA 周次日期区间。")
    base = created.min().normalize()

    errors = []
    for week_idx, col_name, qa_start, qa_end in week_meta:
        expected_start = base + timedelta(days=(week_idx - 1) * 7)
        expected_end = expected_start + timedelta(days=6)
        if qa_start > qa_end:
            errors.append(
                f"{col_name}：开始日期 {qa_start.date()} 晚于结束日期 {qa_end.date()}"
            )
            continue
        if qa_start != expected_start or qa_end != expected_end:
            errors.append(
                f"{col_name}：应为 {expected_start.date()}-{expected_end.date()}，"
                f"实际为 {qa_start.date()}-{qa_end.date()}"
            )

    if errors:
        preview = "; ".join(errors[:5])
        if len(errors) > 5:
            preview += f" ...（另有 {len(errors) - 5} 条）"
        raise AppError(
            "QA 周次日期区间与工单推导周次不一致。 " + preview
        )


def parse_inputs(ticket_file, qa_file) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if ticket_file is None:
        raise AppError("请上传工单明细 Excel。")
    if qa_file is None:
        raise AppError("请上传 QA 结果 Excel。")

    ticket_df = _read_excel(ticket_file)
    qa_raw_df = _remove_qa_example_rows(_read_excel(qa_file))
    qa_df, week_meta = _normalize_qa_df(qa_raw_df)

    ensure_columns(ticket_df, TICKET_REQUIRED_COLUMNS, "工单明细")
    ensure_columns(qa_df, QA_REQUIRED_COLUMNS, "QA结果")
    _validate_qa_week_ranges(ticket_df, week_meta)

    if ticket_df.empty:
        raise AppError("工单明细文件没有数据行。")
    if qa_df.empty:
        raise AppError("QA 结果文件没有数据行。")

    return ticket_df.copy(), qa_df.copy()
