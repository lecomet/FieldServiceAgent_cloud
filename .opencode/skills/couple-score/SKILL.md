---
name: couple-score
description: 装维 CP 积分指标查询。用于用户询问 CP积分、CP积分占比、CP原始积分、CP积分合计、CP人员数等与 CP 积分相关的正式装维人员聚合指标；可按全省、地市、区县过滤。不要用于人员明细、开店/非开店人数或破零率。
---

# Couple Score

## 定位

只负责 CP 积分相关指标。当前口径只算正式人员。

| 指标 | 口径 |
| --- | --- |
| CP人员 | 非开店正式人员：`用工属性` 含 `装维` 或 `智家工程师`，且不是 `装维门店`，且 `新老装维 in 新装维/老装维` |
| CP积分 | 非开店正式人员 `原始积分` 合计 |
| CP积分占比 | CP积分 / 全部正式人员原始积分，不包含实习人员 |

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

"$UV" run --with openpyxl python .opencode/skills/couple-score/scripts/analyze_cp_score.py --official-dir ./temp/data/official --question "装维全省CP积分占比"
"$UV" run --with openpyxl python .opencode/skills/couple-score/scripts/analyze_cp_score.py --official-dir ./temp/data/official --question "聊城CP积分占比" --city "聊城"
"$UV" run --with openpyxl python .opencode/skills/couple-score/scripts/analyze_cp_score.py --official-dir ./temp/data/official --question "东昌府区CP积分合计" --county "东昌府区"
```

## 参数说明

| 参数 | 说明 | 示例 |
| --- | --- | --- |
| `--question` | 用户问题原文，用于记录和识别 | `--question "CP积分占比"` |
| `--city` | 限定地市，可选 | `--city 聊城` |
| `--county` | 限定区县，可选 | `--county 东昌府区` |
| `--official-dir` | 正式人员 Excel 目录 | `--official-dir temp/data/official` |

## 输出说明

固定输出范围、来源文件、样本人数、CP人员数、CP原始积分、全部正式人员原始积分、CP积分占比。

## 约束

- 必须使用 `scripts/analyze_cp_score.py`，不得临时重写 CP 积分计算逻辑。
- 不得临时生成 Python 脚本，不得使用 `Write-File`、`Set-Content`、`echo`、here-string、`python -c` 或临时 `.py` 文件处理数据。
- 不要手动输入中文 Excel 文件名；脚本会按 `temp/data/official/` 查找本轮下载的正式人员 `.xlsx`。
- 所有 Python 脚本文件名保持 ASCII，避免 Windows 终端编码导致脚本路径乱码。
- 不要把实习人员合并进 CP 积分指标。
- “全部人员”在 CP 积分占比中固定指全部正式人员，不包含实习人员。
- 不要把地市汇总或区县汇总表临时当作人员明细使用。
- 查询交付后删除 `temp/data/`。
- 如果用户问人员清单或个人情况，转到 `stuff-metrics`。
- 如果用户问开店人数、非开店人数、开店人员清单，转到 `storefront-staff`。

