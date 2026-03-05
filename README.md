# Customer Service Weekly Performance Automation (DataTagging)

A complete Python + Streamlit project for automated weekly/monthly customer service performance settlement.

## Project Structure

```text
DataTagging/
├── performance_app.py
├── utils/
│   ├── file_parser.py
│   ├── time_calculator.py
│   ├── performance_calc.py
│   ├── exporter.py
│   └── helpers.py
├── requirements.txt
└── README.md
```

## Environment

- Python 3.9+
- Streamlit
- pandas
- openpyxl

## Install

```bash
python3 -m pip install -r requirements.txt
```

## Run

```bash
python3 -m streamlit run performance_app.py
```

## NAS Deployment (Docker)

### 1) Upload project to NAS

Place the whole `DataTagging` folder on your NAS, for example:

```text
/volume1/docker/datatagging/DataTagging
```

### 2) Start service

Run the following in NAS terminal:

```bash
cd /volume1/docker/datatagging/DataTagging
./scripts/nas-start.sh
```

If startup is successful, open:

```text
http://<NAS_IP>:8501
```

### 3) Common operations

View logs:

```bash
./scripts/nas-logs.sh
```

Stop service:

```bash
./scripts/nas-stop.sh
```

## Input Files

### 1) Ticket Detail Excel
Required columns:

- 创建时间
- 关联提出人
- 工单分类
- 工单标签
- 计数
- 预计工时
- 客服结论
- 客服补充

### 2) QA Result Excel
Required columns:

- 客服
- 等级 (A/B/C)
- 所属周次 (W1/W2/...)
- 质检会话占比

## Core Rules Implemented

- Empty or zero `计数` is forced to `1`.
- `预计工时` is converted from hour to minute.
- For regular tickets, if estimated time is missing or zero, default to 10 minutes.
- Week split uses first ticket `创建时间` as baseline, each 7 days is one week (`W1`, `W2`, ...).
- Actual minute calculation and reward ladder are implemented exactly as required.
- Output file is `月度绩效结算表.xlsx` with:
  - Sheet `绩效表（月汇总）`
  - Weekly sheets: `W1`, `W2`, ...

## Error Handling

The app provides explicit errors for:

- Missing uploads
- Missing required columns
- Invalid date parsing
- Missing QA mapping for any agent-week combination
- Excel read failures
