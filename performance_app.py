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


st.set_page_config(page_title="Customer Service Weekly Performance", layout="wide")
st.title("Customer Service Weekly Performance Automation")
st.caption("Upload ticket detail + QA result files, then generate monthly settlement.")

with st.expander("Input format notes", expanded=False):
    st.markdown(
        """
- Ticket detail columns: 创建时间, 关联提出人, 工单分类, 工单标签, 计数, 预计工时, 客服结论, 客服补充
- QA result columns: 客服, 等级, 所属周次, 质检会话占比
- Output file name: 月度绩效结算表.xlsx
"""
    )

ticket_file = st.file_uploader("Upload ticket detail Excel", type=["xlsx"])
qa_template_path = Path(__file__).resolve().parent / "templates" / "qa_template.xlsx"
if qa_template_path.exists():
    st.download_button(
        label="Download QA Template",
        data=qa_template_path.read_bytes(),
        file_name="QA质检模板.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("QA template file not found: templates/qa_template.xlsx")

qa_file = st.file_uploader("Upload QA result Excel", type=["xlsx"])
template_file = st.file_uploader(
    "Upload export template Excel (optional)",
    type=["xlsx"],
    help="If provided, output will follow this template layout.",
)
estimated_unit_label = st.selectbox(
    "Estimated time unit in ticket file",
    options=["Minutes", "Hours"],
    index=0,
)
allow_missing_qa = st.checkbox(
    "Allow missing QA rows (use default grade for missing agent-week)",
    value=False,
)
default_grade = st.selectbox(
    "Default QA grade for missing rows",
    options=["A档", "B档", "C档"],
    index=1,
    disabled=not allow_missing_qa,
)

if st.button("Run Calculation", type="primary"):
    try:
        ticket_df, qa_df = parse_inputs(ticket_file, qa_file)
        estimated_unit = "minutes" if estimated_unit_label == "Minutes" else "hours"
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
        )

        st.success("Calculation completed.")
        if allow_missing_qa and not qa_fallback_df.empty:
            st.warning(
                f"Used default QA grade '{default_grade}' for "
                f"{len(qa_fallback_df)} missing agent-week rows."
            )
            st.dataframe(qa_fallback_df, use_container_width=True)

        st.subheader("Monthly Summary Preview")
        st.dataframe(monthly_df, use_container_width=True)

        st.subheader("Weekly Detail Preview")
        week_names = sorted(weekly_sheets.keys(), key=lambda x: int(x[1:]))
        selected_week = st.selectbox("Select week", options=week_names)
        st.dataframe(weekly_sheets[selected_week], use_container_width=True)

        st.download_button(
            label="Download 月度绩效结算表.xlsx",
            data=excel_bytes,
            file_name="月度绩效结算表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except AppError as exc:
        st.error(str(exc))
    except Exception as exc:  # noqa: BLE001
        st.error(f"Unexpected error: {exc}")
        st.code(traceback.format_exc())
