# FieldServiceAgent BI 下载指南

本目录只服务于 FieldServiceAgent 云端项目；本地调试目录通常为 `D:\FieldServiceAgent_cloud`。

## 定位

- 仅下载 FieldServiceAgent 使用的装维日清单数据。
- 不做分析；分析和查询交给 `stuff-metrics`、`area-metrics`、`city-metrics`、`couple-score`、`storefront-staff`、`first-purchase-rate`。
- 不要为了某个调度目标下载无关数据。

## 下载映射

| 调度目标 | 配置名 | 保存目录 | 文件名模式 |
| --- | --- | --- | --- |
| `stuff-metrics` 正式人员 | `field-service-agent-official-staff` | `temp/data/official/` | `装维随销发展日清单_正式人员_{acct_day}_{month_id}_{timestamp}.xlsx` |
| `stuff-metrics` 实习人员 | `field-service-agent-intern-staff` | `temp/data/intern/` | `装维随销发展日清单_实习人员_{acct_day}_{month_id}_{timestamp}.xlsx` |
| `area-metrics` 地市汇总 | `field-service-agent-area-summary` | `temp/data/area/` | `装维月累计_地市汇总_{acct_day}_{month_id}_{timestamp}.xlsx` |
| `area-metrics` 地市日清单 | `field-service-agent-area-daily` | `temp/data/area/` | `装维日清单_地市汇总_{acct_day}_{month_id}_{timestamp}.xlsx` |
| `city-metrics` 区县汇总 | `field-service-agent-city-summary` | `temp/data/city/` | `装维日清单_区县汇总_{acct_day}_{month_id}_{timestamp}.xlsx` |

人员聚合指标 `couple-score`、`storefront-staff`、`first-purchase-rate` 只需要正式人员明细，即 `field-service-agent-official-staff`。

## 运行要求

云端使用 Linux/bash，先设置 UTF-8：

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi
```

必须传：

```text
month_id=YYYYMM,acct_day=YYYYMMDD
```

## 下载示例

```bash
acctDay="$(python -c "from datetime import date,timedelta; print((date.today()-timedelta(days=1)).strftime('%Y%m%d'))")"
monthId="${acctDay:0:6}"

"$UV" run --with python-dateutil --with requests --with openpyxl python .opencode/skills/bi-data-download/scripts/download_bi_data.py -t config -n field-service-agent-official-staff --params "month_id=${monthId},acct_day=${acctDay}" -o ./temp/data/official

"$UV" run --with python-dateutil --with requests --with openpyxl python .opencode/skills/bi-data-download/scripts/download_bi_data.py -t config -n field-service-agent-intern-staff --params "month_id=${monthId},acct_day=${acctDay}" -o ./temp/data/intern

"$UV" run --with python-dateutil --with requests --with openpyxl python .opencode/skills/bi-data-download/scripts/download_bi_data.py -t config -n field-service-agent-area-summary --params "month_id=${monthId},acct_day=${acctDay}" -o ./temp/data/area

"$UV" run --with python-dateutil --with requests --with openpyxl python .opencode/skills/bi-data-download/scripts/download_bi_data.py -t config -n field-service-agent-area-daily --params "month_id=${monthId},acct_day=${acctDay}" -o ./temp/data/area

"$UV" run --with python-dateutil --with requests --with openpyxl python .opencode/skills/bi-data-download/scripts/download_bi_data.py -t config -n field-service-agent-city-summary --params "month_id=${monthId},acct_day=${acctDay}" -o ./temp/data/city
```

## 约束

- 不检查本地是否已有文件；每次流程都直接下载。
- `area-metrics` 按请求下载 `field-service-agent-area-summary` 或 `field-service-agent-area-daily` 到 `temp/data/area/`，不要下载正式人员明细。
- `city-metrics` 只下载 `field-service-agent-city-summary` 到 `temp/data/city/`，不要下载正式人员明细。
- `stuff-metrics` 查正式人员才下载 `field-service-agent-official-staff`，查实习人员才下载 `field-service-agent-intern-staff`。
- BI 下载格式固定为 `.xlsx`，配置里的 `file_type` 必须是 `xlsx`，下载后必须校验响应内容是标准 `.xlsx`。
- 读取 `.xlsx` 时第 1 行必须是字段名；字段名只清理前后空白、换行和制表符，不维护字段映射。缺字段时报告缺少字段和实际字段。
- 下载前必须校验 `acct_day` 与 `month_id`，未来日期直接拒绝下载。
- 下载文件必须带 `{acct_day}_{month_id}_{timestamp}`，不要用下载时间推断数据日期。
- 云端凭证优先通过 Secret/环境变量处理；没有 Secret/环境变量入口时，只允许使用运行实例内未提交的 `.runtime/.bi_credentials`、`.secrets/bi_credentials` 或 `FSA_BI_CREDENTIALS_FILE` 指向的文件。本地调试可用 `keyring_manager.py` 或 `.bi_credentials`。不要把真实凭证写入文档或上传包。
- 查询、分析或报告交付完成后删除 `temp/data/`；不要保留 BI 下载文件作为长期缓存。

清理命令：

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi
"$UV" run python .opencode/skills/bi-data-download/scripts/cleanup_temp_data.py
```
