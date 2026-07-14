# FieldServiceAgent Cloud Environment

本文档约束云端部署时的运行环境、路径、凭证和企业微信通道使用方式。

## 强制加载

运行 `FieldServiceAgent` 前必须先读取本文件和 `DATA_CONFIG.md`，所有 Agent 和 Skill 操作遵循这里的约束。

## 工作目录

- 云端默认从当前项目根目录运行，不绑定本机盘符。
- 本地调试目录可以是 `D:\FieldServiceAgent_cloud`。
- 云端如需显式指定目录，使用环境变量 `FSA_WORKDIR`。
- 所有脚本路径使用相对项目根目录的 `.opencode/skills/.../scripts/...`。

## 工具路径

云端默认是 Linux/bash。不要在云端使用 PowerShell、`$env:`、`Get-Date`、`Test-Path`、`Get-Command` 或 `& $uv` 这类 Windows 写法。

云端优先使用用户本地工具目录下的 `uv`：

```bash
export PYTHONIOENCODING=utf-8
UV="${HOME}/.local/bin/uv"
if [ ! -x "$UV" ]; then UV="$(command -v uv)"; fi
```

运行 Python 脚本时使用：

```bash
"$UV" run --with openpyxl python .opencode/skills/<skill>/scripts/<script>.py
```

只有在 Windows 本地调试时才使用 PowerShell。

## 临时目录

- BI 下载只写入 `temp/data/...`。
- 查询和报表生成完成后清理 `temp/data/`。
- `output/` 只保存需要交付给用户或通道层的最终文件。
- 云端通道需要持久化文件时，应由外部通道服务上传到对象存储；Skill 不写死对象存储地址。地市图片主动推送直接走企业微信机器人 Webhook。

## 凭证和密钥

云端禁止把以下内容写入仓库文件：

- BI 用户名和密码
- 企业微信 Bot Secret
- 企业微信机器人 Webhook key

推荐优先级：

```text
云端 Secret Manager / 环境变量 > Keyring > 运行时凭证文件 > .bi_credentials（仅本地调试）
```

常用环境变量：

```text
BI_USER
BI_PASS
FSA_BI_CREDENTIALS_FILE
FSA_WECOM_WEBHOOK
FSA_CHROME_BIN
CHROME_BIN
```

如果云端平台没有 Secret/环境变量配置入口，但允许登录运行实例或挂载私有文件，可以只在云端运行环境创建凭证文件，不要放进上传包、仓库或文档。读取顺序如下：

```text
FSA_BI_CREDENTIALS_FILE 指向的文件
.runtime/.bi_credentials
.secrets/bi_credentials
.bi_credentials（仅本地调试）
```

文件格式：

```text
bi_user=...
bi_pass=...
```

`.runtime/`、`.secrets/` 和 `.bi_credentials` 已加入 `.gitignore`。如果平台只能通过上传包提供文件，仍不建议把真实 BI 密码放入上传包；应先让云端平台补 Secret/环境变量或运行时文件能力。

## 图片渲染

地市随销统计图片由 `area-metrics/scripts/generate_and_push_area_report.py` 使用 Chromium/Chrome 截图生成。云端查找顺序包括：

```text
FSA_CHROME_BIN
CHROME_BIN
/usr/local/bin/chromium
/usr/bin/chromium
/usr/bin/chromium-browser
/usr/bin/google-chrome
tools/chrome/chrome-linux64/chrome
```

当前云端如已有 `/usr/local/bin/chromium`，不需要随包安装 Chromium。

云端没有系统中文字体时，随包放置中文字体：

```text
tools/fonts/NotoSansSC-Regular.otf
```

报表 HTML 会自动用 `@font-face` 引用该字体，避免截图中文字变成方块。

## 企业微信通道

本项目不直接维护企业微信长连接。云端 `qili_send_message` 通道层负责文本回复：

```text
企业微信长连接收消息
消息去重
任务排队
调用本 Agent
发送文本
主动推送调度
```

`qili_send_message` 只按字符串文本使用，不作为图片/文件发送通道。Agent 和 Skill 只产出文本、HTML、PNG、Excel 或结构化结果。

明确出现“推送/发送/发到群/企微”且需要发送地市图片报表时，`area-metrics` 调用 `push-sender`，通过企业微信机器人 Webhook 发送图片。

`push-sender` 默认使用固定企业微信机器人 Webhook；如需临时换群，可设置 `FSA_WECOM_WEBHOOK` 覆盖。
