---
name: stuff-metrics
description: 装维人员明细查询。用于用户查看、下载、筛选装维正式人员或实习人员个人明细，按姓名、工号、地市、区县查询个人发展量（移动、终端宽带、FTTR-H、FTTR-B、FTTR-H/B、151、天翼智屏）、积分（发展积分、运营积分、原始积分）和实习人员实习开始时间；数据粒度是一行一个人。不要用于CP积分占比、开店人数、非开店人数、破零率等聚合指标。
---

# Stuff Metrics

## 定位

只负责装维人员明细查询，不负责聚合指标计算。

| 问题类型 | 处理方式 |
| --- | --- |
| 姓名、工号、某个人装维情况 | 使用 `query_detail.py --person` |
| 某些人/某个人的装维随销发展日清单 | 使用 `query_detail.py --person`，取这些人/这个人 |
| 某地市正式/实习人员清单 | 使用 `query_detail.py --city` |
| 某区县正式/实习人员清单 | 使用 `query_detail.py --county` |
| 全省装维随销发展日清单 | 转到 `area-metrics` |
| 某地市/某区县装维随销发展日清单 | 转到 `city-metrics` |
| CP积分占比 | 转到 `couple-score` |
| 开店人数、非开店人数、开店人员 | 转到 `storefront-staff` |
| 破零率 | 转到 `first-purchase-rate` |

## 工作目录

所有命令在当前项目根目录执行；本地调试通常为 `D:\FieldServiceAgent_cloud`，云端使用 `FSA_WORKDIR` 或当前工作目录。

## 数据来源

原始数据由 `bi-data-download` skill 下载：

| 人员类型 | BI 配置名 | 输出目录 |
| --- | --- | --- |
| 正式人员 | `field-service-agent-official-staff` | `temp/data/official/` |
| 实习人员 | `field-service-agent-intern-staff` | `temp/data/intern/` |

下载必须传：`month_id=YYYYMM,acct_day=YYYYMMDD`。每次查询都直接下载本轮数据，不检查或复用历史 Excel。

## 字段范围

人员明细字段来自正式/实习人员日清单：

- 维度字段：`地市`、`区县`、`装维姓名`、`工号`、`用工属性`、`新老装维`、`维护区域`
- 发展量字段：`移动`、`终端宽带`、`FTTR-H`、`FTTR-B`、`FTTR-H/B`、`151`、`天翼智屏`
- 积分字段：`发展积分`、`运营积分`、`原始积分`
- 实习人员扩展字段：`实习开始时间`。该字段只在实习人员 Excel 中出现时输出和保存；脚本按 Excel 实际表头动态保留新增字段。

## 固化执行流程

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi

# 个人查询：不保存，直接输出
"$UV" run --with openpyxl python .opencode/skills/stuff-metrics/scripts/query_detail.py --staff-type official --official-dir ./temp/data/official --intern-dir ./temp/data/intern --person "张三"

# 地市人员明细：保存 Excel，只输出路径和行数
"$UV" run --with openpyxl python .opencode/skills/stuff-metrics/scripts/query_detail.py --staff-type official --official-dir ./temp/data/official --intern-dir ./temp/data/intern --city "聊城"

# 区县人员明细：保存 Excel，20 行以内输出明细，超过 20 行只输出路径和行数
"$UV" run --with openpyxl python .opencode/skills/stuff-metrics/scripts/query_detail.py --staff-type official --official-dir ./temp/data/official --intern-dir ./temp/data/intern --county "东昌府区"
```

清单类问题必须先看匹配行数：20 行以内可以在消息栏罗列；超过 20 行不要罗列明细，只保存 xlsx 并回复文件路径、行数、查询条件。

## 参数说明

| 参数 | 说明 | 示例 |
| --- | --- | --- |
| `--staff-type` | `official` 或 `intern` | `--staff-type official` |
| `--person` | 姓名或工号 | `--person 张三` |
| `--city` | 地市名称 | `--city 聊城` |
| `--county` | 区县名称 | `--county 东昌府区` |
| `--official-dir` | 正式人员 Excel 目录 | `--official-dir temp/data/official` |
| `--intern-dir` | 实习人员 Excel 目录 | `--intern-dir temp/data/intern` |
| `--max-output` | 消息栏最多输出行数，默认 20 | `--max-output 20` |

## 输出说明

- 个人查询：终端直接输出匹配人员记录；如果匹配超过 20 行，要求用户缩小姓名或工号条件。
- 地市查询：保存 `output/city_detail_<地市>_<人员类型>.xlsx`，终端输出路径和行数。
- 区县查询：保存 `output/county_detail_<区县>_<人员类型>.xlsx`，终端输出路径、行数；20 行以内可输出明细。
- 全省查询：保存 `output/province_detail_<人员类型>.xlsx`，终端输出路径和行数。

## 约束

- 必须使用 `scripts/query_detail.py`，不得临时重写人员筛选逻辑。
- 不得临时生成 Python 脚本，不得使用 `Write-File`、`Set-Content`、`echo`、here-string、`python -c` 或临时 `.py` 文件处理数据。
- 不要手动输入中文 Excel 文件名；`query_detail.py` 会按 `temp/data/official/` 或 `temp/data/intern/` 查找本轮下载的 `.xlsx`。
- 所有 Python 脚本文件名保持 ASCII，避免 Windows 终端编码导致脚本路径乱码。
- 读取和输出 Excel 时保持 UTF-8 终端环境。
- 查询前先调用 `bi-data-download` 下载本轮 Excel；查询交付后删除 `temp/data/`。
- 不要在本 skill 中计算 CP积分占比、开店人数、非开店人数、破零率。

