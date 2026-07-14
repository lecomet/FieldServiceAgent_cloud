---
name: storefront-staff
description: 装维开店和非开店人员指标查询。用于用户询问开店人数、非开店人数、开店人员清单、非开店人员清单、开店人员人均积分等正式装维人员指标；可按全省、地市、区县过滤。不要用于 CP积分占比、破零率或普通人员明细。
---

# Storefront Staff

## 定位

只负责开店/非开店人员相关指标。当前口径只算正式人员。

| 分类 | 口径 |
| --- | --- |
| 开店员工 | 正式员工 + `用工属性 = 装维门店` + `新老装维 in 新装维/老装维` |
| 非开店员工 | 正式员工 + `用工属性` 含 `装维` 或 `智家工程师` + `用工属性 != 装维门店` + `新老装维 in 新装维/老装维` |

## 工作目录

所有命令在当前项目根目录执行；本地调试通常为 `D:\FieldServiceAgent_cloud`，云端使用 `FSA_WORKDIR` 或当前工作目录。

## 数据来源

读取本轮下载的正式人员 Excel：

```text
temp/data/official/装维随销发展日清单_正式人员_*.xlsx
```

每次查询前都调用 `bi-data-download` 下载 `field-service-agent-official-staff` 到 `temp/data/official/`，参数为 `month_id=YYYYMM,acct_day=YYYYMMDD`；不要检查或复用历史 Excel。

## 固化执行流程

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi

"$UV" run --with openpyxl python .opencode/skills/storefront-staff/scripts/analyze_storefront_staff.py --official-dir ./temp/data/official --question "全省开店人数有多少"
"$UV" run --with openpyxl python .opencode/skills/storefront-staff/scripts/analyze_storefront_staff.py --official-dir ./temp/data/official --question "聊城非开店人数" --city "聊城"
"$UV" run --with openpyxl python .opencode/skills/storefront-staff/scripts/analyze_storefront_staff.py --official-dir ./temp/data/official --question "东昌府区开店人员清单" --county "东昌府区"
"$UV" run --with openpyxl python .opencode/skills/storefront-staff/scripts/analyze_storefront_staff.py --official-dir ./temp/data/official --question "开店人员中人均积分是多少"
```

## 参数说明

| 参数 | 说明 | 示例 |
| --- | --- | --- |
| `--question` | 用户问题原文，脚本自动识别指标 | `--question "开店人数有多少"` |
| `--metric` | 显式指标，可选：`open_shop_count`、`non_open_shop_count`、`open_shop_avg_score`、`open_shop_list`、`non_open_shop_list` | `--metric open_shop_count` |
| `--city` | 限定地市，可选 | `--city 聊城` |
| `--county` | 限定区县，可选 | `--county 东昌府区` |
| `--official-dir` | 正式人员 Excel 目录 | `--official-dir temp/data/official` |
| `--max-output` | 清单在消息栏最多输出行数，默认 20 | `--max-output 20` |

## 输出说明

- 人数和人均积分：直接输出指标结果、样本人数和来源文件。
- 人员清单：保存到 `output/storefront_<范围>_<分类>.xlsx`。
- 清单超过 20 行时，不在消息栏罗列明细，只回复 Excel 路径和行数。

## 约束

- 必须使用 `scripts/analyze_storefront_staff.py`，不得临时重写开店/非开店计算逻辑。
- 不得临时生成 Python 脚本，不得使用 `Write-File`、`Set-Content`、`echo`、here-string、`python -c` 或临时 `.py` 文件处理数据。
- 不要手动输入中文 Excel 文件名；脚本会按 `temp/data/official/` 查找本轮下载的正式人员 `.xlsx`。
- 所有 Python 脚本文件名保持 ASCII，避免 Windows 终端编码导致脚本路径乱码。
- 不要把实习人员合并进开店/非开店指标。
- 不要省略 `新老装维 in 新装维/老装维` 条件。
- 不要把 CP积分占比放到本 skill，CP 问题转到 `couple-score`。
- 如果用户问个人发展量、个人积分、姓名或工号，转到 `stuff-metrics`。
- 查询交付后删除 `temp/data/`。

