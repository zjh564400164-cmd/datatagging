from __future__ import annotations

import traceback
from pathlib import Path

import pandas as pd
import streamlit as st

from utils.exporter import export_to_excel
from utils.file_parser import parse_inputs
from utils.helpers import AppError
from utils.performance_calc import calculate_performance
from utils.time_calculator import build_week_date_labels, preprocess_and_calculate


st.set_page_config(page_title="客服绩效自动计算", layout="wide")
st.title("客服绩效自动计算")
st.caption("上传工单明细与 QA 质检文件，自动生成月度绩效结算。")

with st.expander("输入格式说明", expanded=False):
    st.markdown(
        """
- 工单明细必需列：创建时间, 关联提出人, 工单分类, 工单标签, 计数, 预计工时, 客服结论, 客服补充
- QA 结果必需列：客服, 等级, 所属周次, 质检会话占比（可选：版本）
- 导出文件名：月度绩效结算表.xlsx
"""
    )

ticket_file = st.file_uploader("上传工单明细 Excel", type=["xlsx"])
qa_template_path = Path(__file__).resolve().parent / "templates" / "qa_template.xlsx"
if qa_template_path.exists():
    st.download_button(
        label="下载 QA 模板",
        data=qa_template_path.read_bytes(),
        file_name="QA质检模板.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("未找到 QA 模板文件：templates/qa_template.xlsx")

qa_file = st.file_uploader("上传 QA 结果 Excel", type=["xlsx"])
export_template_path = (
    Path(__file__).resolve().parent / "templates" / "export_template.xlsx"
)
if export_template_path.exists():
    st.download_button(
        label="下载导出模板",
        data=export_template_path.read_bytes(),
        file_name="导出模板.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("未找到导出模板文件：templates/export_template.xlsx")

template_file = st.file_uploader(
    "上传导出模板 Excel（可选）",
    type=["xlsx"],
    help="上传后将按模板样式导出结果。",
)
estimated_unit_label = st.selectbox(
    "工单文件中的预计工时单位",
    options=["分钟", "小时"],
    index=0,
)
allow_missing_qa = st.checkbox(
    "允许 QA 缺失（缺失客服-周次时使用默认等级）",
    value=False,
)
default_grade = st.selectbox(
    "QA 缺失时默认等级",
    options=["A档", "B档", "C档"],
    index=1,
    disabled=not allow_missing_qa,
)

if st.button("开始计算", type="primary"):
    try:
        ticket_df, qa_df = parse_inputs(ticket_file, qa_file)
        estimated_unit = "minutes" if estimated_unit_label == "分钟" else "hours"
        processed_ticket_df = preprocess_and_calculate(
            ticket_df, estimated_unit=estimated_unit
        )
        week_date_labels = build_week_date_labels(processed_ticket_df)
        monthly_df, weekly_sheets, qa_fallback_df = calculate_performance(
            processed_ticket_df,
            qa_df,
            allow_missing_qa=allow_missing_qa,
            default_grade=default_grade.replace("档", ""),
            default_ratio_percent=0.0,
        )

        template_bytes = template_file.getvalue() if template_file is not None else None
        excel_bytes = export_to_excel(
            monthly_df,
            weekly_sheets,
            template_bytes=template_bytes,
            week_date_labels=week_date_labels,
            ticket_verify_df=ticket_df,
            qa_verify_df=qa_df,
        )

        st.success("计算完成。")
        if allow_missing_qa and not qa_fallback_df.empty:
            st.warning(
                f"已对 {len(qa_fallback_df)} 条缺失的客服-周次记录使用默认 QA 等级「{default_grade}」。"
            )
            st.dataframe(qa_fallback_df, use_container_width=True)

        st.subheader("月度汇总预览")
        st.dataframe(monthly_df, use_container_width=True)

        st.subheader("周明细预览")
        if weekly_sheets:
            week_names = sorted(weekly_sheets.keys(), key=lambda x: int(x[1:]))
            selected_week = st.selectbox("选择周次", options=week_names)
            st.dataframe(weekly_sheets[selected_week], use_container_width=True)
        else:
            st.info("当前结果无周维度明细（例如全部为老版月度计算）。")

        st.download_button(
            label="下载 月度绩效结算表.xlsx",
            data=excel_bytes,
            file_name="月度绩效结算表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except AppError as exc:
        st.error(str(exc))
    except Exception as exc:  # noqa: BLE001
        st.error(f"系统异常：{exc}")
        st.code(traceback.format_exc())
