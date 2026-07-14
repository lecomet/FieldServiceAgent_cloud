#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
【BI数据下载校验模块】文件命名规范、下载触发判断、异常数据预判

功能：
- 文件命名防重复（业务标识 + 日期 + 时间段 + 唯一标识）
- 下载触发判断（按时间戳校验数据是否更新）
- 异常数据预判（早6点数据量异常、全0值等）
- 强制用户交互确认
"""

import os
import re
import time
from datetime import datetime, timedelta
from typing import Optional, Tuple, Dict, Any


class DownloadFileNamer:
    """文件命名规范：业务标识 + 日期 + 时间段 + 唯一标识"""

    BUSINESS_IDS = {
        "sales_monthly": "销售点月报",
        "hr_data": "HR人效",
        "commission": "佣金清单",
        "daily": "电渠日表",
    }

    @classmethod
    def generate_filename(
        cls,
        business_type: str,
        period: str,
        timestamp: str = None,
        extension: str = "xlsx",
    ) -> str:
        """生成标准化的文件名

        格式：{业务标识}_{账期}_{时间段}_{唯一标识}.{扩展名}
        示例：销售点月报_202602_上午_20260213_143052.xlsx
        """
        business_name = cls.BUSINESS_IDS.get(business_type, business_type)
        if timestamp is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{business_name}_{period}_{timestamp}.{extension}"

    @classmethod
    def parse_filename(cls, filename: str) -> Optional[Dict[str, str]]:
        """解析文件名提取关键信息

        返回：{'business_type': 'sales_monthly', 'period': '202602', 'timestamp': '20260213_143052'}
        """
        pattern = r"^(.+?)_(\d{6})_(\d{8}_\d{6})\.(.+)$"
        match = re.match(pattern, filename)
        if not match:
            return None

        business_name, period, timestamp, ext = match.groups()

        reverse_map = {v: k for k, v in cls.BUSINESS_IDS.items()}
        business_type = reverse_map.get(business_name, business_name)

        return {
            "business_type": business_type,
            "period": period,
            "timestamp": timestamp,
            "extension": ext,
            "original": filename,
        }

    @classmethod
    def is_duplicate_download(
        cls, output_dir: str, business_type: str, period: str, hours: int = 24
    ) -> Tuple[bool, Optional[str]]:
        """检查是否存在重复下载（24小时内同名文件）

        返回：(是否重复, 已有文件路径)
        """
        cutoff_time = time.time() - (hours * 3600)
        pattern = f"{cls.BUSINESS_IDS.get(business_type, business_type)}_{period}_"

        if not os.path.exists(output_dir):
            return False, None

        for filename in os.listdir(output_dir):
            if not filename.startswith(pattern) or not filename.endswith(".xlsx"):
                continue
            filepath = os.path.join(output_dir, filename)
            try:
                mtime = os.path.getmtime(filepath)
                if mtime >= cutoff_time:
                    return True, filepath
            except OSError:
                continue

        return False, None


class DownloadTriggerChecker:
    """下载触发判断：按时间戳/时间段校验数据是否更新"""

    @staticmethod
    def get_time_period() -> str:
        """获取当前时间段标识"""
        hour = datetime.now().hour
        if 6 <= hour < 12:
            return "上午"
        elif 12 <= hour < 18:
            return "下午"
        elif 18 <= hour < 24:
            return "晚间"
        else:
            return "凌晨"

    @staticmethod
    def should_skip_download(
        existing_file: str, check_content: bool = True, min_rows: int = 100
    ) -> Tuple[bool, str]:
        """判断是否跳过下载。

        当前策略要求每次直接从 BI 下载，不再因本地文件存在而跳过。

        返回：(是否跳过, 原因)
        """
        return False, "按最新策略每次直接下载，不复用本地文件"

    @staticmethod
    def check_data_update_needed(
        existing_file: str, period: str
    ) -> Tuple[bool, Optional[str]]:
        """检查数据是否需要更新。

        当前策略要求每次直接从 BI 下载。

        返回：(是否需要下载, 原因或已有文件路径)
        """
        return True, "按最新策略每次直接下载，不复用本地文件"


class AbnormalDataPredictor:
    """异常数据预判：早6点数据仅1条、全0值等场景"""

    EARLY_MORNING_HOUR = 6
    MIN_EXPECTED_ROWS = 100
    ZERO_VALUE_THRESHOLD = 0.95

    @classmethod
    def predict_early_morning_anomaly(cls, hour: int = None) -> bool:
        """预判早6点异常：数据量可能不足"""
        if hour is None:
            hour = datetime.now().hour
        return hour < cls.EARLY_MORNING_HOUR

    @classmethod
    def check_row_count_anomaly(cls, row_count: int, hour: int = None) -> Tuple[bool, str]:
        """检查数据行数异常

        返回：(是否异常, 提示信息)
        """
        if hour is None:
            hour = datetime.now().hour

        if hour < cls.EARLY_MORNING_HOUR and row_count < cls.MIN_EXPECTED_ROWS:
            return True, f"[警告] 当前时间 {hour}点，数据仅 {row_count} 行（正常应>{cls.MIN_EXPECTED_ROWS}），可能为昨日数据未汇总"

        if row_count < 10:
            return True, f"[严重] 数据仅 {row_count} 行，疑似异常"

        return False, ""

    @classmethod
    def check_zero_value_anomaly(cls, df) -> Tuple[bool, str]:
        """检查全0值异常

        返回：(是否异常, 提示信息)
        """
        if df is None or len(df) == 0:
            return True, "[严重] 数据为空"

        numeric_cols = df.select_dtypes(include=["number"]).columns
        if len(numeric_cols) == 0:
            return False, ""

        total_cells = len(df) * len(numeric_cols)
        zero_cells = (df[numeric_cols] == 0).sum().sum()

        if total_cells > 0:
            zero_ratio = zero_cells / total_cells
            if zero_ratio >= cls.ZERO_VALUE_THRESHOLD:
                return True, f"[严重] 数据中 {zero_ratio*100:.1f}% 为0值，疑似数据异常"

        return False, ""

    @classmethod
    def validate_and_warn(
        cls, row_count: int, df=None, raise_error: bool = False
    ) -> bool:
        """综合校验并警告

        Args:
            row_count: 数据行数
            df: DataFrame，可选
            raise_error: 是否抛出异常而非仅警告

        Returns:
            是否通过校验（False为异常）

        Raises:
            ValueError: 当 raise_error=True 且检测到异常时
        """
        warnings = []

        if cls.predict_early_morning_anomaly():
            warnings.append(f"[提醒] 当前为早{datetime.now().hour}点，数据可能未完全汇总")

        is_row_anomaly, row_msg = cls.check_row_count_anomaly(row_count)
        if is_row_anomaly:
            warnings.append(row_msg)

        if df is not None:
            is_zero_anomaly, zero_msg = cls.check_zero_value_anomaly(df)
            if is_zero_anomaly:
                warnings.append(zero_msg)

        if warnings:
            print("\n" + "=" * 50)
            print("[警告] 数据异常检测：")
            for w in warnings:
                print(f"  {w}")
            print("=" * 50)

            if raise_error:
                raise ValueError("; ".join(warnings))
            return False

        return True


def prompt_user_confirmation(message: str, default: bool = False) -> bool:
    """强制用户交互确认（禁止模型自行决策）

    返回：用户确认结果
    """
    print()
    print("=" * 50)
    print(f"[确认] {message}")
    print("=" * 50)

    suffix = "[Y/n]" if default else "[y/N]"
    while True:
        try:
            response = input(f"请确认是否继续? {suffix}: ").strip().lower()
            if response == "":
                return default
            if response in ("y", "yes"):
                return True
            if response in ("n", "no"):
                return False
            print("请输入 y 或 n")
        except KeyboardInterrupt:
            print("\n[取消] 用户中断")
            return False


def require_user_confirmation_for_anomaly(
    row_count: int, df=None, message: str = "检测到数据异常，是否继续?"
) -> bool:
    """异常数据必须触发用户交互确认

    返回：用户确认是否继续
    """
    cls = AbnormalDataPredictor

    is_anomaly = False
    warnings = []

    if cls.predict_early_morning_anomaly():
        warnings.append(f"当前为早{datetime.now().hour}点")
        is_anomaly = True

    is_row_anomaly, row_msg = cls.check_row_count_anomaly(row_count)
    if is_row_anomaly:
        warnings.append(row_msg)
        is_anomaly = True

    if df is not None:
        is_zero_anomaly, zero_msg = cls.check_zero_value_anomaly(df)
        if is_zero_anomaly:
            warnings.append(zero_msg)
            is_anomaly = True

    if is_anomaly:
        full_message = f"{message}\n异常: {'; '.join(warnings)}"
        return prompt_user_confirmation(full_message, default=False)

    return True
