from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Tuple

import pandas as pd

from .helpers import AppError, normalize_text, safe_to_float

STANDARD_WEEKLY_MINUTES = 2400.0


@dataclass
class WeeklyPerformance:
    agent_name: str
    week: str
    grade: str
    quality_ratio: float
    x_factor: float
    y_actual: float
    n_standard: float
    z_rate: float
    m_over: float
    reward: float


def calc_x_factor(grade: str, ratio_percent: float) -> float:
    grade = normalize_text(grade).upper()
    # Support both 0-1 ratio and 0-100 percent inputs.
    if ratio_percent <= 1:
        ratio_percent = ratio_percent * 100
    if grade == "A":
        if ratio_percent < 20:
            return 1.0
        if ratio_percent < 50:
            return 1.1
        return 1.2
    if grade == "B":
        return 0.85
    if grade == "C":
        return 0.6
    raise AppError(f"无效的 QA 等级：{grade}")


def calc_week_reward(grade: str, m_over: float) -> float:
    grade = normalize_text(grade).upper()

    if m_over <= 0:
        base_reward = 0.0
    elif m_over <= 240:
        base_reward = 25.0
    elif m_over < 480:
        base_reward = 100.0
    elif m_over < 720:
        base_reward = 175.0
    elif m_over < 960:
        base_reward = 250.0
    elif m_over < 1200:
        base_reward = 325.0
    elif m_over < 1440:
        base_reward = 400.0
    else:
        base_reward = 400.0

    if grade == "C":
        return 0.0
    if grade == "B":
        return min(base_reward, 250.0)
    if grade == "A":
        return min(base_reward, 400.0)

    raise AppError(f"无效的 QA 等级：{grade}")


def _prepare_qa_map(qa_df: pd.DataFrame) -> Dict[Tuple[str, str], dict]:
    qa_map: Dict[Tuple[str, str], dict] = {}
    for _, row in qa_df.iterrows():
        agent_name = normalize_text(row.get("客服"))
        week = normalize_text(row.get("所属周次"))
        grade = normalize_text(row.get("等级")).upper()
        ratio_percent = safe_to_float(row.get("质检会话占比"), 0.0)

        if not agent_name or not week:
            continue

        qa_map[(agent_name, week)] = {
            "grade": grade,
            "ratio_percent": ratio_percent,
        }
    return qa_map


def calculate_performance(
    ticket_df: pd.DataFrame,
    qa_df: pd.DataFrame,
    allow_missing_qa: bool = False,
    default_grade: str = "B",
    default_ratio_percent: float = 0.0,
):
    if ticket_df.empty:
        raise AppError("没有可用于计算绩效的工单数据。")

    grouped = (
        ticket_df.groupby(["关联提出人", "week"], as_index=False)
        .agg(
            total_tickets=("count_fixed", "sum"),
            y_actual=("actual_minutes", "sum"),
        )
        .rename(columns={"关联提出人": "agent_name"})
    )

    qa_map = _prepare_qa_map(qa_df)

    missing_pairs = []
    for _, row in grouped.iterrows():
        agent_name = normalize_text(row["agent_name"])
        week = normalize_text(row["week"])
        if (agent_name, week) not in qa_map:
            missing_pairs.append((agent_name, week))

    if missing_pairs and not allow_missing_qa:
        preview = ", ".join([f"{a}-{w}" for a, w in missing_pairs[:10]])
        if len(missing_pairs) > 10:
            preview += f" ...（另有 {len(missing_pairs) - 10} 条）"
        raise AppError(
            "以下客服-周次缺少 QA 数据："
            f"{preview}。请补齐所有周次的 QA 等级。"
        )

    qa_fallback_used: List[Tuple[str, str]] = []
    weekly_rows = []
    for _, row in grouped.iterrows():
        agent_name = normalize_text(row["agent_name"])
        week = normalize_text(row["week"])
        qa_item = qa_map.get((agent_name, week))
        if qa_item is None:
            qa_item = {
                "grade": normalize_text(default_grade).upper(),
                "ratio_percent": safe_to_float(default_ratio_percent, 0.0),
            }
            qa_fallback_used.append((agent_name, week))

        grade = qa_item["grade"]
        ratio_percent = qa_item["ratio_percent"]
        x = calc_x_factor(grade, ratio_percent)
        y = safe_to_float(row["y_actual"], 0.0)
        n = STANDARD_WEEKLY_MINUTES
        z = (x * y) / n if n > 0 else 0.0
        m = (x * y) - n
        reward = calc_week_reward(grade, m)

        weekly_rows.append(
            WeeklyPerformance(
                agent_name=agent_name,
                week=week,
                grade=grade,
                quality_ratio=ratio_percent,
                x_factor=x,
                y_actual=y,
                n_standard=n,
                z_rate=z,
                m_over=m,
                reward=reward,
            )
        )

    weekly_df = pd.DataFrame(
        [
            {
                "客服姓名": r.agent_name,
                "质检等级": r.grade,
                "质检系数 X": r.x_factor,
                "周实际工时 Y": r.y_actual,
                "周标准工时 N": r.n_standard,
                "绩效达成率 Z": r.z_rate,
                "超标准时间 M": r.m_over,
                "周奖励": r.reward,
                "week": r.week,
            }
            for r in weekly_rows
        ]
    )

    total_tickets_df = grouped.groupby("agent_name", as_index=False)["total_tickets"].sum()
    total_tickets_df = total_tickets_df.rename(
        columns={"agent_name": "客服姓名", "total_tickets": "总工单量"}
    )

    month_perf = (
        weekly_df.assign(corrected=lambda d: d["质检系数 X"] * d["周实际工时 Y"])
        .groupby("客服姓名", as_index=False)
        .agg(
            累计修正工时=("corrected", "sum"),
            累计激励奖金=("周奖励", "sum"),
            n_sum=("周标准工时 N", "sum"),
        )
    )
    month_perf["月度达成率"] = month_perf.apply(
        lambda r: (r["累计修正工时"] / r["n_sum"]) if r["n_sum"] > 0 else 0.0, axis=1
    )

    monthly_df = month_perf.drop(columns=["n_sum"])
    monthly_df = monthly_df.merge(total_tickets_df, on="客服姓名", how="left")

    monthly_df = monthly_df[
        ["客服姓名", "总工单量", "累计修正工时", "累计激励奖金", "月度达成率"]
    ].sort_values("客服姓名")

    weekly_sheets = {}
    for week_name in sorted(weekly_df["week"].unique(), key=lambda w: int(w[1:])):
        sheet_df = weekly_df[weekly_df["week"] == week_name].copy()
        sheet_df = sheet_df[
            [
                "客服姓名",
                "质检等级",
                "质检系数 X",
                "周实际工时 Y",
                "周标准工时 N",
                "绩效达成率 Z",
                "超标准时间 M",
                "周奖励",
            ]
        ].sort_values("客服姓名")
        weekly_sheets[week_name] = sheet_df

    qa_fallback_df = pd.DataFrame(
        [{"客服姓名": a, "所属周次": w} for a, w in qa_fallback_used]
    )
    return monthly_df.reset_index(drop=True), weekly_sheets, qa_fallback_df
