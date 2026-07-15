---
name: monthly-bigscreen
description: 月累计联动大屏。用于用户明确要求生成、制作、导出、展示或推送“大屏/可视化大屏/酷炫大屏/柱状图大屏”时，结合各地市装维月累计、各地市装维日清单和正式人员装维月累计生成 HTML/PNG，并可按明确推送词发送企业微信。普通各地市随销月累计报表仍由 area-metrics 处理。
---

# Monthly Bigscreen

## 定位

只负责“月累计联动大屏”：地市月累计 + 地市日清单 + 正式人员月累计的可视化大屏。

不要处理普通月累计报表；用户只说“各地市随销月累计报表/各地市随销月清单报表”时，转 `area-metrics` 使用 `generate_area_html_report.py --report-kind monthly`。

## 触发

使用本 skill 的条件：

- 用户明确包含“大屏/可视化/酷炫/联动/柱状图大屏”。
- 或用户明确说“生成大屏/制作大屏/导出大屏/推送大屏/发送大屏”。
- 如同时出现“个人+地市月累计/正式人员+地市/月累计个人榜”，也使用本 skill。

不使用本 skill 的条件：

- “各地市随销月累计报表”“各地市随销月清单报表”：走 `area-metrics` 普通报表。
- “各地市随销月清单/各地市随销月累计清单”且无大屏或报表生成词：只查询回复。

## 数据

| 数据 | BI 配置名 | 保存目录 | 文件名 |
| --- | --- | --- | --- |
| 各地市装维月累计 | `field-service-agent-area-summary` | `temp/data/area/` | `装维月累计_地市汇总_{acct_day}_{month_id}_{timestamp}.xlsx` |
| 各地市装维日清单 | `field-service-agent-area-daily` | `temp/data/area/` | `装维日清单_地市汇总_{acct_day}_{month_id}_{timestamp}.xlsx` |
| 正式人员装维月累计 | 暂未配置固定 BI 名称时使用本地已下载文件 | `temp/data/` 或显式路径 | `正式人员装维月累计*.xlsx` |

如果缺少地市日清单或正式人员月累计 Excel，直接说明缺少文件；不要伪造 BI 路径，不用正式人员日清单替代月累计。

## 生成 HTML

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi

"$UV" run --with openpyxl python .opencode/skills/monthly-bigscreen/scripts/generate_monthly_bigscreen_report.py \
  --city-file ./temp/data/area/装维月累计_地市汇总_${acctDay}_${monthId}_${timestamp}.xlsx \
  --daily-city-file ./temp/data/area/装维日清单_地市汇总_${acctDay}_${monthId}_${timestamp}.xlsx \
  --staff-file ./temp/data/正式人员装维月累计.xlsx \
  --acct-day "${acctDay}" \
  --month-id "${monthId}"
```

也可以不传 `--city-file/--staff-file`，脚本会从 `temp/data/` 自动查找最新匹配文件。

## 生成 PNG 或推送

用户只说生成/制作/导出大屏：生成 HTML 和 PNG，不推送。

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi

"$UV" run --with openpyxl python .opencode/skills/monthly-bigscreen/scripts/generate_and_push_monthly_bigscreen.py \
  --acct-day "${acctDay}" \
  --month-id "${monthId}"
```

用户明确说“推送/发送/发到群/企微 + 大屏”时，才加 `--send`：

```bash
"$UV" run --with openpyxl --with requests python .opencode/skills/monthly-bigscreen/scripts/generate_and_push_monthly_bigscreen.py \
  --acct-day "${acctDay}" \
  --month-id "${monthId}" \
  --send
```

推送走 `push-sender` 的企业微信内部网关配置；不要在回复中展示完整 Webhook key。

## 输出

默认输出到：

- `output/各地市随销月累计大屏_{acct_day}.html`
- `output/各地市随销月累计大屏_{acct_day}.png`

默认截图尺寸为 `1920x1400`，用于完整容纳月累计、当日地市表现和个人榜。

完成后报告使用的地市文件、人员文件、输出路径、地市行数和人员行数。只清理 `.part`、`.invalid` 等失败/中间文件，保留已校验 `.xlsx` 和最终输出。
