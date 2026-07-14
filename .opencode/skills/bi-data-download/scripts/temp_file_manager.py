#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
【统一临时文件管理模块】提供临时文件创建、清理、生命周期管理

功能：
- 临时文件/目录创建（统一 temp_ 前缀命名）
- 自动清理机制（Skill 执行完成后）
- 留存文件管理（按周期留存，超期自动清理）
"""

import os
import shutil
import time
from datetime import datetime, timedelta
from typing import Optional, List
from contextlib import contextmanager


TEMP_PREFIX = "temp_"
DEFAULT_RETENTION_DAYS = 7


class TempFileManager:
    def __init__(self, base_dir: str = None):
        self.base_dir = base_dir or os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "temp")
        self.created_files: List[str] = []

    def create_temp_file(self, prefix: str = "", suffix: str = "") -> str:
        """创建临时文件，返回文件路径"""
        os.makedirs(self.base_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pid = os.getpid()
        name = f"{TEMP_PREFIX}{prefix}{timestamp}_{pid}{suffix}"
        filepath = os.path.join(self.base_dir, name)
        self.created_files.append(filepath)
        return filepath

    def create_temp_dir(self, prefix: str = "") -> str:
        """创建临时目录，返回目录路径"""
        os.makedirs(self.base_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pid = os.getpid()
        name = f"{TEMP_PREFIX}{prefix}{timestamp}_{pid}"
        dirpath = os.path.join(self.base_dir, name)
        os.makedirs(dirpath, exist_ok=True)
        self.created_files.append(dirpath)
        return dirpath

    def cleanup(self):
        """清理本次创建的临时文件/目录"""
        for path in self.created_files:
            try:
                if os.path.isfile(path):
                    os.remove(path)
                elif os.path.isdir(path):
                    shutil.rmtree(path)
            except Exception as e:
                print(f"[警告] 清理失败: {path}, {e}")
        self.created_files.clear()

    @staticmethod
    def cleanup_all_temp(base_dir: str = None, retention_days: int = DEFAULT_RETENTION_DAYS):
        """清理所有超期临时文件"""
        if base_dir is None:
            base_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "temp")

        if not os.path.exists(base_dir):
            return

        cutoff_time = time.time() - (retention_days * 24 * 60 * 60)
        removed_count = 0

        for name in os.listdir(base_dir):
            if not name.startswith(TEMP_PREFIX):
                continue
            path = os.path.join(base_dir, name)
            try:
                mtime = os.path.getmtime(path)
                if mtime < cutoff_time:
                    if os.path.isfile(path):
                        os.remove(path)
                    elif os.path.isdir(path):
                        shutil.rmtree(path)
                    removed_count += 1
            except Exception as e:
                print(f"[警告] 清理失败: {path}, {e}")

        if removed_count > 0:
            print(f"[清理] 已清理 {removed_count} 个超期临时文件")


class OutputFileManager:
    def __init__(self, output_dir: str = None, retention_days: int = 30):
        self.output_dir = output_dir or os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "output")
        self.retention_days = retention_days

    def cleanup_old_files(self):
        """清理超期结果文件"""
        if not os.path.exists(self.output_dir):
            return

        cutoff_time = time.time() - (self.retention_days * 24 * 60 * 60)
        removed_count = 0

        for name in os.listdir(self.output_dir):
            if name.startswith(TEMP_PREFIX):
                continue
            path = os.path.join(self.output_dir, name)
            try:
                mtime = os.path.getmtime(path)
                if mtime < cutoff_time:
                    if os.path.isfile(path):
                        os.remove(path)
                    elif os.path.isdir(path):
                        shutil.rmtree(path)
                    removed_count += 1
            except Exception as e:
                print(f"[警告] 清理失败: {path}, {e}")

        if removed_count > 0:
            print(f"[清理] 已清理 {removed_count} 个超期结果文件")


@contextmanager
def temp_file_context(prefix: str = "", suffix: str = "", cleanup_on_exit: bool = True):
    """临时文件上下文管理器"""
    manager = TempFileManager()
    filepath = manager.create_temp_file(prefix, suffix)
    try:
        yield filepath
    finally:
        if cleanup_on_exit:
            manager.cleanup()


def ensure_directory(path: str, is_temp: bool = False) -> str:
    """确保目录存在"""
    if is_temp:
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "temp", path)
    os.makedirs(path, exist_ok=True)
    return path