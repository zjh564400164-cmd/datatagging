from __future__ import annotations

from datetime import datetime
from typing import Any, Iterable

import pandas as pd


class AppError(Exception):
    """Custom application-level exception for user-facing errors."""


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def is_empty(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def safe_to_float(value: Any, default: float = 0.0) -> float:
    if is_empty(value):
        return default
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def parse_datetime(value: Any, field_name: str) -> datetime:
    if is_empty(value):
        raise AppError(f"字段「{field_name}」存在空日期时间值。")
    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        raise AppError(f"字段「{field_name}」包含无效日期时间值。")
    return dt.to_pydatetime()


def ensure_columns(df: pd.DataFrame, required_cols: Iterable[str], sheet_name: str) -> None:
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise AppError(
            f"{sheet_name}缺少必填列：{', '.join(missing)}"
        )
