#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import base64
import hashlib
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from generate_monthly_bigscreen_report import main as generate_main

MAX_WECOM_IMAGE_SIZE = 2 * 1024 * 1024
DEFAULT_WECOM_GATEWAY_URL = "http://137.0.249.205:9600/serviceAgent/rest/qywx/qwgroup/send"
DEFAULT_WECOM_APP_ID = "41e806212a70360b59356834e3e21335"
DEFAULT_WECOM_APP_KEY = "6d64aed4f7973cdd21c5f3a28e69f564"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


def project_root() -> Path:
    return Path(__file__).resolve().parents[4]


def find_chrome() -> Path:
    root = project_root()
    candidates = [
        os.getenv("FSA_CHROME_BIN", "").strip(),
        os.getenv("CHROME_BIN", "").strip(),
        shutil.which("chromium"),
        shutil.which("chromium-browser"),
        shutil.which("google-chrome"),
        shutil.which("google-chrome-stable"),
        shutil.which("chrome"),
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/usr/local/bin/chromium",
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
        root / "tools" / "chrome" / "chrome-linux64" / "chrome",
        root / "tools" / "chrome" / "chrome",
    ]
    for candidate in candidates:
        if not candidate:
            continue
        path = Path(candidate).expanduser()
        if path.exists():
            return path.resolve()
    raise FileNotFoundError("未找到 Chrome/Chromium，无法把大屏 HTML 渲染为图片。")


def render_png(html_path: Path, image_path: Path, width: int, height: int) -> None:
    chrome = find_chrome()
    image_path.parent.mkdir(parents=True, exist_ok=True)
    profile_dir = image_path.parent / ".chrome-bigscreen-profile"
    profile_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        str(chrome),
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-extensions",
        "--hide-scrollbars",
        "--no-first-run",
        "--no-default-browser-check",
        "--allow-file-access-from-files",
        "--run-all-compositor-stages-before-draw",
        "--virtual-time-budget=1000",
        f"--user-data-dir={profile_dir}",
        f"--window-size={width},{height}",
        f"--screenshot={image_path}",
        html_path.as_uri(),
    ]
    try:
        subprocess.run(cmd, check=True, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=90)
    except subprocess.TimeoutExpired:
        if image_path.exists() and image_path.stat().st_size > 0:
            return
        raise


def extract_webhook_key(webhook: str) -> str:
    parsed = urlparse(webhook)
    key = parse_qs(parsed.query).get("key", [""])[0].strip()
    if key:
        return key
    if webhook and "://" not in webhook and "?" not in webhook:
        return webhook.strip()
    raise ValueError("企业微信 Webhook 未包含 key 参数。")


def build_wecom_gateway_request(webhook: str) -> tuple[str, dict[str, str]]:
    key = extract_webhook_key(webhook)
    gateway_url = os.getenv("FSA_WECOM_GATEWAY_URL", DEFAULT_WECOM_GATEWAY_URL).strip()
    app_id = os.getenv("FSA_WECOM_APP_ID", DEFAULT_WECOM_APP_ID).strip()
    app_key = os.getenv("FSA_WECOM_APP_KEY", DEFAULT_WECOM_APP_KEY).strip()
    if not app_id or not app_key:
        raise ValueError("企业微信内部网关缺少 X-APP-ID 或 X-APP-KEY。")
    return (
        f"{gateway_url}?key={key}",
        {
            "Content-Type": "application/json",
            "X-APP-ID": app_id,
            "X-APP-KEY": app_key,
        },
    )


def push_image(root: Path, image_path: Path) -> bool:
    config_path = root / ".opencode" / "skills" / "push-sender" / "push_config.json"
    if not config_path.exists():
        raise FileNotFoundError(f"推送配置不存在: {config_path}")
    config = json.loads(config_path.read_text(encoding="utf-8"))
    webhook = os.getenv("FSA_WECOM_WEBHOOK", "").strip() or config.get("wecom_webhook")
    if not webhook:
        raise ValueError("企业微信 Webhook 未配置。")

    data = image_path.read_bytes()
    if len(data) > MAX_WECOM_IMAGE_SIZE:
        raise ValueError(f"图片超过企业微信机器人 2MB 限制: {len(data)} bytes")

    import requests

    payload = {
        "msgtype": "image",
        "image": {
            "base64": base64.b64encode(data).decode("ascii"),
            "md5": hashlib.md5(data).hexdigest(),
        },
    }
    url, headers = build_wecom_gateway_request(webhook)
    resp = requests.post(url, json=payload, headers=headers, timeout=30)
    resp.raise_for_status()
    result = resp.json()
    print(f"企业微信返回: errcode={result.get('errcode')} errmsg={result.get('errmsg')}")
    return result.get("errcode") == 0


def build_generate_argv(args: argparse.Namespace) -> list[str]:
    argv = ["generate_monthly_bigscreen_report.py"]
    for flag, value in [
        ("--city-file", args.city_file),
        ("--daily-city-file", args.daily_city_file),
        ("--staff-file", args.staff_file),
        ("--acct-day", args.acct_day),
        ("--month-id", args.month_id),
        ("--output", args.html_output),
    ]:
        if value:
            argv.extend([flag, value])
    return argv


def resolve_html_path(root: Path, args: argparse.Namespace) -> Path:
    if args.html_output:
        html_path = Path(args.html_output)
        if not html_path.is_absolute():
            html_path = root / html_path
        return html_path
    if args.acct_day:
        return root / "output" / f"各地市随销月累计大屏_{args.acct_day}.html"
    candidates = [
        path
        for path in (root / "output").glob("各地市随销月累计大屏_*.html")
        if path.is_file()
    ]
    if not candidates:
        raise FileNotFoundError("未找到刚生成的月累计大屏 HTML。")
    return max(candidates, key=lambda path: path.stat().st_mtime)


def main() -> None:
    parser = argparse.ArgumentParser(description="生成月累计联动大屏 HTML/PNG，并可推送企业微信")
    parser.add_argument("--city-file", help="地市月累计 Excel；默认从 temp/data 自动查找")
    parser.add_argument("--daily-city-file", help="各地市装维日清单 Excel；默认从 temp/data 自动查找")
    parser.add_argument("--staff-file", help="正式人员月累计 Excel；默认从 temp/data 自动查找")
    parser.add_argument("--acct-day", help="截至日期 YYYYMMDD")
    parser.add_argument("--month-id", help="账期 YYYYMM")
    parser.add_argument("--html-output", help="HTML 输出路径")
    parser.add_argument("--image-output", help="PNG 输出路径；默认与 HTML 同名")
    parser.add_argument("--width", type=int, default=1920, help="图片渲染宽度")
    parser.add_argument("--height", type=int, default=1400, help="图片渲染高度")
    parser.add_argument("--send", action="store_true", help="生成图片后推送企业微信")
    args = parser.parse_args()

    root = project_root()
    old_argv = sys.argv[:]
    try:
        sys.argv = build_generate_argv(args)
        generate_main()
    finally:
        sys.argv = old_argv

    html_path = resolve_html_path(root, args)

    if args.image_output:
        image_path = Path(args.image_output)
        if not image_path.is_absolute():
            image_path = root / image_path
    else:
        image_path = html_path.with_suffix(".png")

    render_png(html_path, image_path, args.width, args.height)
    print(f"图片文件: {image_path}")
    print(f"图片大小: {image_path.stat().st_size} bytes")

    if args.send:
        if not push_image(root, image_path):
            raise SystemExit("企业微信推送失败。")
        print("推送结果: 已发送到企业微信内部网关")
    else:
        print("推送结果: 未发送；传 --send 后发送到企业微信机器人Webhook")


if __name__ == "__main__":
    main()
