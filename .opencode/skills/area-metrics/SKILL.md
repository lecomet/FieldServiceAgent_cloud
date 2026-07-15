---
name: area-metrics
description: 装维地市汇总查询。用于用户查看全省、山东省、各地市装维随销发展日清单、地市随销统计、各地市随销月清单、生成各地市随销月累计普通报表或地市汇总，数据粒度是一行一个地市，字段包含装维发展量、全量发展量、占比、积分、人数、人均等。月累计大屏/可视化大屏应转到 monthly-bigscreen；某个地市的装维随销发展日清单应转到 city-metrics 取该地市所有区县。
---

# Area Metrics

## 定位

只负责全省/山东省地市汇总信息。不要处理人员明细或区县汇总。

| 数据 | BI 配置名 | 输出目录 |
| --- | --- | --- |
| 各地市装维月累计 | `field-service-agent-area-summary` | `temp/data/area/` |

必传参数：`month_id=YYYYMM,acct_day=YYYYMMDD`。

每次查询或报表流程都调用 `bi-data-download` 下载 `field-service-agent-area-summary` 并保存到 `temp/data/area/`；下载脚本会优先复用同 `acct_day`、`month_id` 且通过保存校验的本地文件。不要下载 `field-service-agent-official-staff` 或 `field-service-agent-intern-staff`。

## 脚本

必须使用固定脚本查询和格式化地市汇总：

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi

"$UV" run --with openpyxl python .opencode/skills/area-metrics/scripts/query_area_summary.py --area-dir ./temp/data/area
```

脚本会：

- 读取 `temp/data/area/` 本轮下载 Excel。
- 将地市名称标准化：去掉末尾 `市`，例如 `泰安` 和 `泰安市` 统一为 `泰安`。
- 合并同一地市重复行。
- `装维占比` 类字段按合并后的分子/分母重算，不直接相加。
- 宽表拆成多个窄表输出。
- 默认不另存地市汇总副本；原始下载文件只保留在 `temp/data/area/` 到本轮流程结束。只有显式传 `--output` 时才保存格式化副本。

## 数据源

`省公司数据集/大数据和AI运营中心/WXY/美好家/各地市装维月累计`

## 适用问题

- `给我全省装维日清单`
- `给我全省的装维随销发展日清单`
- `给我山东省所有地市装维随销发展日清单`
- `给我各地市装维月累计`
- `各地市随销月清单`
- `截止到2026年7月14日各地市随销月累计清单`
- `截止到2026年7月14日各地市随销月累计报表`
- `地市随销统计`
- `各地市随销统计`
- `生成地市随销统计报表`
- `生成地市随销统计报表并发送`
- `给我各地市20270706的随销日清单`
- `给我推送各地市20270706的随销日清单`
- `各地市装维发展量和积分`

## 规则

- 地市汇总数据保存到 `temp/data/area/`。
- 文件名类似 `装维月累计_地市汇总_20260702_202607_20260703160602.xlsx` 的地市汇总文件必须保存到 `temp/data/area/`。
- 必须使用 `scripts/query_area_summary.py`，不要临时写 Python 脚本合并地市。
- 本 skill 读取 `temp/data/area/` 的本轮下载文件；不要把地市汇总下载位置写成 `output/`。
- 下载文件必须是标准 `.xlsx`，第 1 行必须是字段名；读取时只清理字段名前后空白、换行和制表符，不维护字段映射。
- 缺字段时直接报告缺少字段和实际字段；如果只读取到“全省”而没有各地市明细，直接报告地市维度缺失。
- 用户说“地市随销统计”“各地市随销统计”“各地市日清单”“各地市随销月清单”或“各地市随销月累计清单”且没有“报表/大屏/生成/制作/出一份/导出/页面/推送/发送/发到群/企微”时，直接下载或复用 `field-service-agent-area-summary`，在聊天框展示结果，不生成报表，不推送。
- 用户说“各地市随销月累计报表”“各地市随销月清单报表”，或同时包含“月累计/月清单”和“报表/生成/制作/出一份/导出/页面”时，下载或复用 `field-service-agent-area-summary` 到 `temp/data/area/`，默认文件名类似 `装维月累计_地市汇总_20260714_202607_20260715144516.xlsx`，然后使用 `scripts/generate_area_html_report.py --report-kind monthly` 生成本地 HTML 报表，不推送。
- 用户明确说要“生成/制作/导出/推送大屏”“可视化大屏”“酷炫大屏”“联动大屏”，或“结合个人和地市月累计/个人+地市/正式人员+地市”做大屏时，转到 `monthly-bigscreen`；本 skill 不生成联动大屏。
- 用户说“截止到YYYY年M月D日...”时，将该日期解析为 `acct_day=YYYYMMDD`，`month_id=YYYYMM`；例如“截止到2026年7月14日各地市随销月累计报表”使用 `acct_day=20260714,month_id=202607`。
- 用户说“生成地市随销统计报表”但没有推送词时，使用 `scripts/generate_and_push_area_report.py` 生成本地 HTML、PNG，不推送企业微信；HTML 上方必须先展示所有业务量宽表，下方再放 `151装维量`、`FTTR-H/B装维量`、`人均价值积分` 三项 TOP 5 和 LAST 5 柱状图。
- 用户明确说“推送”“发送”“发到群”或“企微”时，使用 `scripts/generate_and_push_area_report.py --send` 生成 HTML、PNG，并调用 `push-sender` 通过企业微信内部网关推送图片；推送图片同样保持“上方宽表、下方三项柱状图”的版式。
- `推送昨天各地市随销日清单`、`推送昨天各地市装维月累计`、`发送昨天全省地市随销统计` 必须走图片推送流程：下载 `field-service-agent-area-summary` 到 `temp/data/area/`，再执行 `scripts/generate_and_push_area_report.py --send`。不要只用 `qili_send_message` 发送文本摘要。
- 如果积分字段为 0 或暂未同步，仍然生成并推送图片；图片中如实展示 0 或空值，不得因此退化为文本推送。
- 全省/山东省装维随销发展日清单走本 skill，输出山东省所有地市。
- 某个地市的装维随销发展日清单不要走本 skill，转到 `city-metrics` 取该地市所有区县。
- 地市问题不要使用 `stuff-metrics` 的人员明细数据聚合。
- 如果用户问某个人，转到 `stuff-metrics`。
- 如果用户问区县，转到 `city-metrics`。

## 下载示例

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi
acctDay="$(python -c "from datetime import date,timedelta; print((date.today()-timedelta(days=1)).strftime('%Y%m%d'))")"
monthId="${acctDay:0:6}"
"$UV" run --with python-dateutil --with requests --with openpyxl python .opencode/skills/bi-data-download/scripts/download_bi_data.py -t config -n field-service-agent-area-summary --params "month_id=${monthId},acct_day=${acctDay}" -o ./temp/data/area
```

## 报表生成和推送示例

生成各地市随销月累计 HTML 报表，不推送：

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi
"$UV" run --with openpyxl python .opencode/skills/area-metrics/scripts/generate_area_html_report.py --area-dir ./temp/data/area --report-kind monthly
```

生成 HTML 和 PNG，不推送：

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi
"$UV" run --with openpyxl python .opencode/skills/area-metrics/scripts/generate_and_push_area_report.py --area-dir ./temp/data/area
```

生成地市随销统计报表并通过企业微信机器人 Webhook 推送到企业微信（仅用户明确说推送/发送/发到群/企微时）：

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi
"$UV" run --with openpyxl --with requests python .opencode/skills/area-metrics/scripts/generate_and_push_area_report.py --area-dir ./temp/data/area --send
```

输出默认保存到 `output/地市随销统计_{acct_day}.html` 和 `output/地市随销统计_{acct_day}.png`。图片推送走 `push-sender` 的企业微信内部网关；默认 key 从 Webhook 配置里提取，可用 `FSA_WECOM_WEBHOOK` 临时覆盖。

查询、报表生成或推送完成后，只清理 `.part`、`.invalid` 等失败/中间文件；通过保存校验的 `.xlsx` 保留在 `temp/data/area/` 供同日复用。最终 HTML/PNG 保留在 `output/`。

