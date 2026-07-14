# FieldServiceAgent Cloud Data Rules

本文档约束 FieldServiceAgent 云端查询、下载、校验和推送数据的规则。

## 数据下载

- 每次用户发起查询、分析、报表或推送流程，都直接从 BI 下载本轮数据。
- 不复用 `data/` 或旧的 `temp/data/` 文件作为业务结果。
- 下载目录按调度目标固定：

| 调度目标 | BI 配置 | 下载目录 |
| --- | --- | --- |
| `area-metrics` | `field-service-agent-area-summary` | `temp/data/area/` |
| `city-metrics` | `field-service-agent-city-summary` | `temp/data/city/` |
| `stuff-metrics` 正式人员 | `field-service-agent-official-staff` | `temp/data/official/` |
| `stuff-metrics` 实习人员 | `field-service-agent-intern-staff` | `temp/data/intern/` |
| `couple-score` / `storefront-staff` / `first-purchase-rate` | `field-service-agent-official-staff` | `temp/data/official/` |

## Excel 格式

- BI 下载格式固定为 `.xlsx`。
- 第 1 行必须是字段名。
- 字段名只清理前后空白、换行和制表符。
- 不维护字段映射，不猜测别名。
- 缺字段时必须返回缺失字段和实际字段。

## 日期校验

所有 FieldServiceAgent BI 下载必须传：

```text
month_id=YYYYMM
acct_day=YYYYMMDD
```

下载前必须校验：

- `acct_day` 必须是 8 位日期。
- `month_id` 必须是 6 位月份。
- `month_id` 必须等于 `acct_day` 的前 6 位。
- `acct_day` 不能大于当前日期。

示例：

```text
20260706 -> month_id=202607, acct_day=20260706
20270706 -> 如果当前日期早于 2027-07-06，直接拒绝下载
```

## 数据质量校验

下载成功不等于业务成功。脚本必须继续校验：

- 地市汇总必须包含非“全省”的地市行。
- 区县汇总必须包含区县明细。
- 人员明细不能只有表头。
- 报表图片生成失败时不能继续推送。

## 企业微信输出

- 不含“推送/发送/发到群/企微”的问题，只在聊天框展示结果或生成本地文件。
- 明确包含推送词时，才生成图片并交给 `push-sender` 通过企业微信机器人 Webhook 发送。
- `推送昨天各地市随销日清单` 这类请求必须生成图片并走 Webhook；长连接文本只能作为回执，不能替代图片推送。
- 主动推送由云端任务调度触发，不依赖用户聊天上下文。
