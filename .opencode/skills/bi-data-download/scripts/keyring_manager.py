#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
【统一密码管理模块】提供 Keyring 规范的密码存储与读取

功能：
- 安全存储账号密码到系统密钥链
- 跨 Skill 密码复用
- 密码配置交互提醒
"""

import keyring
import getpass
import os
import sys
from typing import Optional, Tuple

SERVICE_NAME = "BI_Data_System"
BI_USER_KEY = f"{SERVICE_NAME}.bi_user"
BI_PASS_KEY = f"{SERVICE_NAME}.bi_pass"


def get_password_from_keyring(key: str) -> Optional[str]:
    """从系统密钥链获取密码"""
    try:
        return keyring.get_password(SERVICE_NAME, key)
    except Exception:
        return None


def set_password_to_keyring(key: str, value: str) -> bool:
    """设置密码到系统密钥链"""
    try:
        keyring.set_password(SERVICE_NAME, key, value)
        return True
    except Exception as e:
        print(f"[错误] 密码存储失败: {e}")
        return False


def load_credentials_interactive() -> Tuple[Optional[str], Optional[str]]:
    """交互式输入凭证（安全提示）"""
    print("\n" + "=" * 50)
    print("【安全提示】请在下方输入您的凭证")
    print("用户名和密码输入时都不会显示")
    print("=" * 50)
    print()

    username = getpass.getpass("请输入BI系统用户名: ")
    if not username:
        print("[失败] 用户名不能为空")
        return None, None

    password = getpass.getpass("请输入BI系统密码: ")
    if not password:
        print("[失败] 密码不能为空")
        return None, None

    return username, password


def load_credentials_keyring() -> Tuple[Optional[str], Optional[str]]:
    """从 Keyring 加载凭证"""
    username = get_password_from_keyring(BI_USER_KEY)
    password = get_password_from_keyring(BI_PASS_KEY)
    return username, password


def save_credentials_keyring(username: str, password: str) -> bool:
    """保存凭证到 Keyring"""
    success_user = set_password_to_keyring(BI_USER_KEY, username)
    success_pass = set_password_to_keyring(BI_PASS_KEY, password)
    return success_user and success_pass


def load_credentials(
    credentials_file: str = None,
    allow_interactive: bool = True,
    auto_save_keyring: bool = False
) -> Tuple[Optional[str], Optional[str]]:
    """统一凭证加载接口

    加载优先级：
    1. 环境变量（BI_USER/BI_PASS 或 FSA_BI_USER/FSA_BI_PASS）
    2. Keyring（系统密钥链）
    3. 配置文件（.bi_credentials）
    4. 交互式输入（如果 allow_interactive=True）

    Args:
        credentials_file: 配置文件路径
        allow_interactive: 允许交互式输入
        auto_save_keyring: 是否自动保存到 Keyring

    Returns:
        (username, password) 元组
    """
    username, password = None, None

    if credentials_file is None:
        credential_files = [
            os.getenv("FSA_BI_CREDENTIALS_FILE"),
            ".runtime/.bi_credentials",
            ".secrets/bi_credentials",
            ".bi_credentials",
        ]
    else:
        credential_files = [credentials_file]

    username = (os.getenv("BI_USER") or os.getenv("FSA_BI_USER") or "").strip()
    password = (os.getenv("BI_PASS") or os.getenv("FSA_BI_PASS") or "").strip()

    if not username or not password:
        username, password = load_credentials_keyring()

    if not username or not password:
        for candidate in [path for path in credential_files if path]:
            if os.path.exists(candidate):
                username, password = _load_from_file(candidate)
                if username and password:
                    break

    if not username or not password:
        if allow_interactive:
            username, password = load_credentials_interactive()
            if username and password and auto_save_keyring:
                if save_credentials_keyring(username, password):
                    print("[提示] 凭证已保存到系统密钥链")

    return username, password


def _load_from_file(filepath: str) -> Tuple[Optional[str], Optional[str]]:
    """从文件加载凭证（已废弃，仅作兼容）"""
    import os
    username, password = None, None
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip()
                    if key == "bi_user":
                        username = value
                    elif key == "bi_pass":
                        password = value
    except Exception:
        pass
    return username, password


def check_credentials_configured() -> bool:
    """检查凭证是否已配置"""
    username, password = load_credentials_keyring()
    return bool(username and password)


def clear_credentials_keyring() -> bool:
    """清除 Keyring 中的凭证"""
    try:
        keyring.delete_password(SERVICE_NAME, BI_USER_KEY)
        keyring.delete_password(SERVICE_NAME, BI_PASS_KEY)
        return True
    except Exception:
        return False
