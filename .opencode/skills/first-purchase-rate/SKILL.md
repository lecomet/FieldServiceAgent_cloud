---
name: first-purchase-rate
description: 装维破零率指标查询。用于用户询问正式装维人员移动、终端宽带、FTTR-H、FTTR-B、FTTR-H/B、151、天翼智屏或任一产品的发展破零率、破零人数、未破零人数；可按全省、地市、区县过滤。不要用于人员明细、CP积分或开店人员指标。
---

# First Purchase Rate

## 定位

只负责正式装维人员破零率。当前脚本口径：

```text
破零人数 = 指定产品发展量 > 0 的正式人员数
破零率 = 破零人数 / 匹配范围内正式人员数
```

如果用户没有指定产品，则使用任一发展量字段大于 0 作为“综合破零”。

## 工作目录

所有命令在当前项目根目录执行；本地调试通常为 `D:\FieldServiceAgent_cloud`，云端使用 `FSA_WORKDIR` 或当前工作目录。

## 数据来源

读取本轮下载的正式人员 Excel：

```text
temp/data/official/装维随销发展日清单_正式人员_*.xlsx
```

每次查询前都调用 `bi-data-download` 下载 `field-service-agent-official-staff` 到 `temp/data/official/`，参数为 `month_id=YYYYMM,acct_day=YYYYMMDD`；不要检查或复用历史 Excel。

## 支持产品

| 用户说法 | 使用字段 |
| --- | --- |
| 移动 | `移动` |
| 宽带、终端宽带 | `终端宽带` |
| FTTR-H | `FTTR-H` |
| FTTR-B | `FTTR-B` |
| FTTR、FTTR-H/B | `FTTR-H/B` |
| 151 | `151` |
| 天翼智屏、智屏 | `天翼智屏` |

## 固化执行流程

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi

"$UV" run --with openpyxl python .opencode/skills/first-purchase-rate/scripts/analyze_first_purchase_rate.py --official-dir ./temp/data/official --question "全省装维破零率"
"$UV" run --with openpyxl python .opencode/skills/first-purchase-rate/scripts/analyze_first_purchase_rate.py --official-dir ./temp/data/official --question "聊城FTTR破零率" --city "聊城"
"$UV" run --with openpyxl python .opencode/skills/first-purchase-rate/scripts/analyze_first_purchase_rate.py --official-dir ./temp/data/official --question "东昌府区151破零率" --county "东昌府区"
"$UV" run --with openpyxl python .opencode/skills/first-purchase-rate/scripts/analyze_first_purchase_rate.py --official-dir ./temp/data/official --product "天翼智屏"
```

## 参数说明

| 参数 | 说明 | 示例 |
| --- | --- | --- |
| `--question` | 用户问题原文，脚本自动识别产品 | `--question "FTTR破零率"` |
| `--product` | 显式产品字段，可选 | `--product "151"` |
| `--city` | 限定地市，可选 | `--city 聊城` |
| `--county` | 限定区县，可选 | `--county 东昌府区` |
| `--official-dir` | 正式人员 Excel 目录 | `--official-dir temp/data/official` |

## 输出说明

固定输出范围、来源文件、产品字段、样本人数、破零人数、未破零人数、破零率。

## 约束

- 必须使用 `scripts/analyze_first_purchase_rate.py`，不得临时重写破零率计算逻辑。
- 不得临时生成 Python 脚本，不得使用 `Write-File`、`Set-Content`、`echo`、here-string、`python -c` 或临时 `.py` 文件处理数据。
- 不要手动输入中文 Excel 文件名；脚本会按 `temp/data/official/` 查找本轮下载的正式人员 `.xlsx`。
- 所有 Python 脚本文件名保持 ASCII，避免 Windows 终端编码导致脚本路径乱码。
- 只算正式人员。
- 如果业务口径要求分母排除无效人员或按人力表计算，先说明当前脚本口径，不要静默改口径。
- 如果用户问某个人发展量，转到 `stuff-metrics`。
- 查询交付后删除 `temp/data/`。

