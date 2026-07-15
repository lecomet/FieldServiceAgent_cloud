---
name: bi-data-download
description: FieldServiceAgent 专用 BI 下载工具，用于从永洪 BI 下载装维正式人员明细、实习人员明细、地市汇总和区县汇总 Excel。用于用户要求下载、更新、刷新或准备装维相关 BI 数据时；不维护无关 BI 指标。
---

# BI Data Download

## Scope

This skill is only for FieldServiceAgent. Use it to download 装维人员明细、地市汇总和区县汇总数据 from BI.

Do not add or mention unrelated BI metrics in this skill.

## Working Directory

Run commands from:

```text
项目根目录；本地调试通常为 D:\FieldServiceAgent_cloud，云端使用 FSA_WORKDIR 或当前工作目录。
```

## Configs

| Data | Config | Required parameters | Output directory |
| --- | --- | --- | --- |
| 装维随销发展日清单（正式人员） | `field-service-agent-official-staff` | `month_id=YYYYMM,acct_day=YYYYMMDD` | `temp/data/official/` |
| 装维随销发展日清单（实习人员） | `field-service-agent-intern-staff` | `month_id=YYYYMM,acct_day=YYYYMMDD` | `temp/data/intern/` |
| 各地市装维月累计 | `field-service-agent-area-summary` | `month_id=YYYYMM,acct_day=YYYYMMDD` | `temp/data/area/` |
| 各地市装维日清单 | `field-service-agent-area-daily` | `month_id=YYYYMM,acct_day=YYYYMMDD` | `temp/data/area/` |
| 各区县装维日清单 | `field-service-agent-city-summary` | `month_id=YYYYMM,acct_day=YYYYMMDD` | `temp/data/city/` |

The config definitions are in `download_configs.json`.

## Output filename convention

All FieldServiceAgent BI downloads must include the data day, month, and download timestamp in the filename:

```text
<业务名>_{acct_day}_{month_id}_{timestamp}.xlsx
```

Example:

```text
装维日清单_区县汇总_20260703_202607_20260703142418.xlsx
```

Meanings:

- `acct_day`: BI query day, from `acct_day=YYYYMMDD`.
- `month_id`: BI query month, from `month_id=YYYYMM`.
- `timestamp`: actual download time, format `yyyyMMddHHmmss`.

Do not infer the data date from file modified time or download time alone. Use `acct_day` and `month_id` from the filename or BI params.

## Commands

Cloud runs on Linux/bash. Set UTF-8 output and locate `uv` from PATH:

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi
acctDay="$(python -c "from datetime import date,timedelta; print((date.today()-timedelta(days=1)).strftime('%Y%m%d'))")"
monthId="${acctDay:0:6}"
```

Download formal staff:

```bash
"$UV" run --with python-dateutil --with requests --with openpyxl python .opencode/skills/bi-data-download/scripts/download_bi_data.py -t config -n field-service-agent-official-staff --params "month_id=${monthId},acct_day=${acctDay}" -o ./temp/data/official
```

Download intern staff:

```bash
"$UV" run --with python-dateutil --with requests --with openpyxl python .opencode/skills/bi-data-download/scripts/download_bi_data.py -t config -n field-service-agent-intern-staff --params "month_id=${monthId},acct_day=${acctDay}" -o ./temp/data/intern
```

Download area summary:

```bash
"$UV" run --with python-dateutil --with requests --with openpyxl python .opencode/skills/bi-data-download/scripts/download_bi_data.py -t config -n field-service-agent-area-summary --params "month_id=${monthId},acct_day=${acctDay}" -o ./temp/data/area
```

Download area daily summary:

```bash
"$UV" run --with python-dateutil --with requests --with openpyxl python .opencode/skills/bi-data-download/scripts/download_bi_data.py -t config -n field-service-agent-area-daily --params "month_id=${monthId},acct_day=${acctDay}" -o ./temp/data/area
```

Download city/county summary:

```bash
"$UV" run --with python-dateutil --with requests --with openpyxl python .opencode/skills/bi-data-download/scripts/download_bi_data.py -t config -n field-service-agent-city-summary --params "month_id=${monthId},acct_day=${acctDay}" -o ./temp/data/city
```

## Rules

- Always pass `month_id=YYYYMM` and `acct_day=YYYYMMDD`.
- Before downloading, check whether a local file for the same `acct_day` and `month_id` already exists and passes saved-file validation. Reuse valid local files; do not repeatedly download BI data for the same day.
- Use `--force` only when the user explicitly asks to refresh/re-download or when the local file fails validation.
- If the user provides neither date nor month, use yesterday for `acct_day` and current month for `month_id`.
- If the user provides a date only, use that date for `acct_day` and that date's month for `month_id`.
- If the user provides both date and month, use both values exactly as requested.
- 下载前必须校验：`acct_day` 是 8 位日期，`month_id` 是 6 位月份，`month_id = acct_day[:6]`，且 `acct_day` 不大于当前日期。
- Save formal staff files to `temp/data/official/`.
- Save intern staff files to `temp/data/intern/`.
- Save area summary files to `temp/data/area/`.
- Save city/county summary files to `temp/data/city/`.
- When the target skill is `area-metrics`, download only the requested area file (`field-service-agent-area-summary` for month cumulative, `field-service-agent-area-daily` for city daily) to `temp/data/area/`; do not download formal staff data.
- When the target skill is `city-metrics`, download only `field-service-agent-city-summary` to `temp/data/city/`; do not download formal staff data.
- Do not delete downloaded `.xlsx` files that pass saved-file validation. Keep them under `temp/data/...` for same-day reuse.
- Output filenames must use `{acct_day}_{month_id}_{timestamp}`.
- BI 下载格式固定为 `.xlsx`，配置里的 `file_type` 必须使用 `xlsx`，不要使用 `excel`、`csv` 或把其他格式改名为 `.xlsx`。
- 下载后必须先保存为 `.part`，校验通过后再改名为正式 `.xlsx`；如果 BI 返回 XML/HTML/CSV/错误页，直接报错并停止保存。
- 保存校验必须确认文件存在、大小大于 0、可由 `openpyxl` 打开、第 1 行有字段名、至少有 1 行有效数据。
- 地市汇总保存校验必须包含 `地市` 字段，且不能只有 `全省` 汇总，必须至少有一个具体地市。
- 区县汇总保存校验必须包含 `地市` 和 `区县` 字段，且 `区县` 非空的数据行数必须大于 0。
- Read downloaded `.xlsx` files with `openpyxl`; do not treat Excel files as text/CSV or guess GBK/ANSI decoding.
- 第 1 行必须是字段名；读取时只清理字段名前后空白、换行和制表符，不维护字段映射。
- 缺字段时直接报告缺少字段和实际字段。
- Do not store BI usernames or passwords in this skill.
- Do not perform analysis here; use `stuff-metrics`, `area-metrics`, or `city-metrics` for querying.

## Cleanup

After delivering the query, analysis, or report result, run this only to remove failed/intermediate downloads such as `.part` and `.invalid` files:

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi
"$UV" run python .opencode/skills/bi-data-download/scripts/cleanup_temp_data.py
```

This keeps validated `.xlsx` files for reuse. To delete all BI temp data, including valid `.xlsx` files, run with `--delete-valid` only when explicitly requested.
