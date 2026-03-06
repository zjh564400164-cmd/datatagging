from __future__ import annotations

from datetime import datetime

import pandas as pd

from .helpers import AppError, is_empty, normalize_text, parse_datetime, safe_to_float

DEFAULT_ESTIMATED_MINUTES = 10.0
PROMOTION_KEYWORDS = ["每日推广数据上传", "每日推广数据上创", "每日推广数据"]
TRAFFIC_KEYWORDS = ["每日流量上传", "每日流量"]


def parse_estimated_minutes(value, estimated_unit: str = "minutes") -> float:
    """
    Convert estimated value to minutes.
    If empty or <=0, returns 0. The caller decides default behavior.
    """
    raw = safe_to_float(value, default=0.0)
    if raw <= 0:
        return 0.0
    if estimated_unit == "hours":
        return raw * 60.0
    return raw


def _detect_internal_upload_type(question_name: str) -> str:
    text = normalize_text(question_name)
    if not text:
        return ""
    if any(k in text for k in PROMOTION_KEYWORDS):
        return "promotion"
    if any(k in text for k in TRAFFIC_KEYWORDS):
        return "traffic"
    return ""


def _internal_upload_minutes(upload_type: str, monthly_count: float) -> float:
    if upload_type == "promotion":
        if monthly_count <= 100:
            return 10.0
        if monthly_count <= 300:
            return 7.0
        return 5.0
    if upload_type == "traffic":
        if monthly_count <= 1000:
            return 1.4
        if monthly_count <= 2000:
            return 1.0
        return 0.8
    return 0.0


def _apply_internal_upload_estimate_rules(df: pd.DataFrame) -> pd.DataFrame:
    if "问题名称" not in df.columns:
        return df

    out = df.copy()
    out["internal_upload_type"] = out["问题名称"].apply(_detect_internal_upload_type)
    matched = out["internal_upload_type"] != ""
    if not matched.any():
        return out

    out["month_key"] = out["创建时间"].dt.to_period("M").astype(str)

    grouped = (
        out.loc[matched]
        .groupby(["关联提出人", "month_key", "internal_upload_type"], as_index=False)
        .agg(monthly_count=("count_fixed", "sum"))
    )
    grouped["override_minutes"] = grouped.apply(
        lambda r: _internal_upload_minutes(
            str(r["internal_upload_type"]), safe_to_float(r["monthly_count"], 0.0)
        ),
        axis=1,
    )
    merge_cols = ["关联提出人", "month_key", "internal_upload_type"]
    out = out.merge(grouped[merge_cols + ["override_minutes"]], on=merge_cols, how="left")

    override_mask = matched & out["override_minutes"].notna() & (out["override_minutes"] > 0)
    out.loc[override_mask, "estimated_minutes"] = out.loc[override_mask, "override_minutes"]

    out = out.drop(columns=["month_key", "override_minutes"])
    return out


def preprocess_and_calculate(ticket_df: pd.DataFrame, estimated_unit: str = "minutes") -> pd.DataFrame:
    df = ticket_df.copy()

    # Parse datetime and sort to guarantee first record is well-defined.
    try:
        df["创建时间"] = pd.to_datetime(df["创建时间"], errors="coerce")
    except Exception as exc:  # noqa: BLE001
        raise AppError(f"解析「创建时间」失败：{exc}") from exc

    if df["创建时间"].isna().any():
        raise AppError("字段「创建时间」包含无效值。")

    df = df.sort_values("创建时间").reset_index(drop=True)

    # Count preprocessing.
    df["count_fixed"] = df["计数"].apply(lambda x: 1 if safe_to_float(x, 0.0) <= 0 else safe_to_float(x, 1.0))

    # Estimated minutes conversion.
    df["estimated_minutes"] = df["预计工时"].apply(
        lambda x: parse_estimated_minutes(x, estimated_unit=estimated_unit)
    )
    df = _apply_internal_upload_estimate_rules(df)

    # Week assignment by first ticket date.
    base_time: datetime = parse_datetime(df.loc[0, "创建时间"], "创建时间")
    day_offset = (df["创建时间"] - pd.Timestamp(base_time)).dt.days
    week_index = (day_offset // 7) + 1
    df["week"] = "W" + week_index.astype(int).astype(str)

    df["actual_minutes"] = df.apply(_calc_actual_minutes, axis=1)
    return df


def build_week_date_labels(processed_df: pd.DataFrame) -> dict[str, str]:
    labels = {}
    grouped = processed_df.groupby("week", as_index=False).agg(
        start_date=("创建时间", "min"),
        end_date=("创建时间", "max"),
    )
    for _, row in grouped.iterrows():
        week = str(row["week"])
        start_dt = pd.to_datetime(row["start_date"])
        end_dt = pd.to_datetime(row["end_date"])
        labels[week] = f"{start_dt.month}月{start_dt.day}日-{end_dt.month}月{end_dt.day}日"
    return labels


def _calc_actual_minutes(row: pd.Series) -> float:
    conclusion = normalize_text(row.get("客服结论"))
    category = normalize_text(row.get("工单分类"))
    tag = normalize_text(row.get("工单标签"))
    supplement = normalize_text(row.get("客服补充"))

    count = safe_to_float(row.get("count_fixed"), 1.0)
    estimated = safe_to_float(row.get("estimated_minutes"), 0.0)
    estimated_or_default = estimated if estimated > 0 else DEFAULT_ESTIMATED_MINUTES

    # Regular ticket.
    if conclusion == "客服可业务处理":
        return count * estimated_or_default

    is_escalation = "转给" in conclusion

    # Escalation - reconciliation submit ticket always fixed 3.
    reconciliation_submit_keywords = ["提交对账工单", "提交对账"]
    if is_escalation and any(keyword in tag for keyword in reconciliation_submit_keywords):
        return 3.0

    # Escalation - investigation.
    if is_escalation and ("问题排查" in category) and not is_empty(supplement):
        return count * estimated_or_default * 0.7

    # Escalation - all others.
    if is_escalation:
        return 3.0

    # Fallback for non-regular and non-escalation rows.
    return count * estimated_or_default
