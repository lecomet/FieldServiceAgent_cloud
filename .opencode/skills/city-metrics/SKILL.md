---
name: city-metrics
description: 装维区县汇总查询和地市日通报 HTML。用于用户查看某个地市的装维随销发展日清单（取该地市所有区县）、某个区县的装维随销发展日清单、区县随销统计、各区县或区县维度装维发展汇总，或生成某个地市日通报 HTML；数据粒度是一行一个区县，字段包含装维发展量、全量发展量、占比、积分、人数、人均等。
---

# City Metrics

## 定位

只负责区县汇总信息和地市日通报 HTML。某个地市的装维随销发展日清单也走本 skill，取该地市所有区县。不要处理人员明细。

| 数据 | BI 配置名 | 输出目录 |
| --- | --- | --- |
| 各区县装维日清单 | `field-service-agent-city-summary` | `temp/data/city/` |

必传参数：`month_id=YYYYMM,acct_day=YYYYMMDD`。

## 脚本

必须使用固定脚本查询和格式化区县汇总：

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi

# 某地市：取该地市所有区县
"$UV" run --with openpyxl python .opencode/skills/city-metrics/scripts/query_city_summary.py --city-dir ./temp/data/city --city "泰安"

# 某区县：取该区县
"$UV" run --with openpyxl python .opencode/skills/city-metrics/scripts/query_city_summary.py --city-dir ./temp/data/city --county "东昌府区"
```

脚本会：

- 读取 `temp/data/city/` 本轮下载 Excel。
- 将地市名称标准化：去掉末尾 `市`，例如 `泰安` 和 `泰安市` 统一为 `泰安`。
- 合并同一地市同一区县重复行。
- `装维占比` 按合并后的 `装维发展量 / 全量发展量` 重算，不直接相加。
- 宽表拆成多个窄表输出。
- 默认不另存区县汇总副本；原始下载文件只保留在 `temp/data/city/` 到本轮流程结束。只有显式传 `--output` 时才保存格式化副本。

生成某个地市日通报 HTML：

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi

"$UV" run --with openpyxl python .opencode/skills/city-metrics/scripts/generate_city_daily_html_report.py --city "聊城" --city-dir ./temp/data/city
```

地市日通报脚本会：

- 读取 `temp/data/city/` 本轮下载 Excel。
- 筛选指定地市下所有区县。
- 生成 `output/地市日通报_{地市}_{acct_day}.html`。
- HTML 上方先展示该地市所有区县的业务量宽表，下方放 `151`、`FTTR-H/B`、`人均价值积分/人均原始积分/原始积分` 可用字段的 TOP 5 和 LAST 5 柱状图。
- 只生成 HTML，不推送企业微信，不调用 `push-sender`。

## 数据源

`省公司数据集/大数据和AI运营中心/WXY/美好家/各区县装维日清单`

## 适用问题

- `给我各区县装维日清单`
- `给我聊城的装维随销发展日清单`
- `给我东昌府区的装维随销发展日清单`
- `看一下东昌府区装维发展汇总`
- `区县随销统计`
- `各区县随销统计`
- `聊城各区县装维发展量`
- `生成聊城昨天装维日通报`
- `聊城地市日通报`
- `给我聊城20260708的装维日通报`

## 规则

- 区县汇总数据保存到 `temp/data/city/`。
- 必须使用 `scripts/query_city_summary.py`，不要临时写 Python 脚本合并地市或区县。
- 本 skill 读取 `temp/data/city/` 的本轮下载文件；不要把区县汇总下载位置写成 `output/`。
- 用户说“区县随销统计”或“各区县随销统计”时，直接下载并读取 `field-service-agent-city-summary` 对应的各区县装维日清单。
- 某个地市的装维随销发展日清单走本 skill，过滤该地市，返回该地市所有区县。
- 某个区县的装维随销发展日清单走本 skill，过滤该区县，返回该区县。
- 用户说“生成/制作/查看 + 某地市 + 日通报/地市日通报/装维日通报”时，下载 `field-service-agent-city-summary` 到 `temp/data/city/`，然后使用 `scripts/generate_city_daily_html_report.py --city <地市>` 生成 HTML。
- 区县只做查询，不生成区县 HTML 报表，不推送。
- 地市日通报只生成 HTML，不推送；即使用户写“推送地市日通报”，当前也先返回 HTML 路径并说明地市日通报未接推送。
- 全省/山东省装维随销发展日清单不要走本 skill，转到 `area-metrics`。
- 区县问题不要使用 `stuff-metrics` 的人员明细数据聚合。
- 如果用户问某个人，转到 `stuff-metrics`。
- 如果用户问全省/山东省，转到 `area-metrics`。

## 下载示例

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi
acctDay="$(python -c "from datetime import date,timedelta; print((date.today()-timedelta(days=1)).strftime('%Y%m%d'))")"
monthId="${acctDay:0:6}"
"$UV" run --with python-dateutil --with requests --with openpyxl python .opencode/skills/bi-data-download/scripts/download_bi_data.py -t config -n field-service-agent-city-summary --params "month_id=${monthId},acct_day=${acctDay}" -o ./temp/data/city
```

查询完成后，删除 `temp/data/city/`。如需保存格式化副本，显式使用 `--output` 写入 `output/`。

## 地市日通报生成示例

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi
acctDay="$(python -c "from datetime import date,timedelta; print((date.today()-timedelta(days=1)).strftime('%Y%m%d'))")"
monthId="${acctDay:0:6}"
"$UV" run --with python-dateutil --with requests --with openpyxl python .opencode/skills/bi-data-download/scripts/download_bi_data.py -t config -n field-service-agent-city-summary --params "month_id=${monthId},acct_day=${acctDay}" -o ./temp/data/city
"$UV" run --with openpyxl python .opencode/skills/city-metrics/scripts/generate_city_daily_html_report.py --city "聊城" --city-dir ./temp/data/city
```

