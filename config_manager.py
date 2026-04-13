# -*- coding: utf-8 -*-
"""Configuration management for New Customer Operation."""

from __future__ import annotations

import math
import os

import pandas as pd


class ConfigManager:
    """Manage config_customer.csv."""

    CONFIG_FILE_NAME = "config_customer.csv"

    REQUIRED_FIELDS = [
        "Files_Import_URL",
        "Environment",
        "Test_API_Base_URL",
        "Account",
        "Password",
    ]

    # A = key, B = value, C = remark.
    DEFAULT_TEMPLATE = [
        ["基础信息", "内容", "备注"],
        ["Files_Import_URL", "", "客户文件导入路径"],
        ["Environment", "test", "运行环境: test 或 prod"],
        ["Test_API_Base_URL", "http://10.103.9.184:7850", "测试环境 API 根地址"],
        ["Prod_API_Base_URL", "", "生产环境 API 根地址"],
        ["Account", "", "ODM 登录账号"],
        ["Password", "", "ODM 登录密码"],
    ]

    def __init__(self):
        self.desktop_url = os.path.join(os.path.expanduser("~"), "Desktop")
        self.config_dir = os.path.join(self.desktop_url, "config")
        self.config_path = os.path.join(self.config_dir, self.CONFIG_FILE_NAME)

    def get_config(self) -> dict:
        """Get config, creating a default file when missing."""
        if not os.path.exists(self.config_path):
            self.create_default_config()
        return self.read_config()

    def read_config(self) -> dict:
        """Read config_customer.csv into a key/value dict."""
        if not os.path.exists(self.config_path):
            raise FileNotFoundError(f"配置文件不存在: {self.config_path}")
        return {row[0]: row[1] for row in self._read_config_rows()}

    def create_default_config(self) -> None:
        """Create the default config file, overwriting an existing file."""
        self._ensure_config_dir()
        self._write_rows(self.DEFAULT_TEMPLATE)

    def get_sync_preview(self) -> dict:
        """Return whether config needs sync and prompt text for UI."""
        if not os.path.exists(self.config_path):
            return {
                "needs_sync": True,
                "reason": "missing",
                "message": (
                    "未检测到配置文件，是否现在生成默认配置文件？\n\n"
                    f"路径：{self.config_path}"
                ),
            }

        existing_rows = self._read_config_rows()
        merged_rows = self._build_synced_rows(existing_rows)
        if existing_rows == merged_rows:
            return {
                "needs_sync": False,
                "reason": "up_to_date",
                "message": "",
            }

        return {
            "needs_sync": True,
            "reason": "template_changed",
            "message": (
                "检测到配置文件与当前代码模板不一致，是否现在更新？\n\n"
                "已填写的配置值会保留，字段结构、顺序和备注会同步到最新模板。\n\n"
                f"路径：{self.config_path}"
            ),
        }

    def sync_config(self) -> bool:
        """Sync the config file to the current template."""
        preview = self.get_sync_preview()
        if not preview["needs_sync"]:
            return False

        if preview["reason"] == "missing":
            self.create_default_config()
            return True

        self._write_rows(self._build_synced_rows(self._read_config_rows()))
        return True

    def set_config_value(self, key: str, value: str) -> None:
        """Update one config value while preserving the template structure."""
        preview = self.get_sync_preview()
        if preview["needs_sync"]:
            self.sync_config()

        rows = self._read_config_rows()
        updated = False
        for row in rows:
            if row[0] == key:
                row[1] = value
                updated = True
                break

        if not updated:
            rows.append([key, value, ""])

        self._write_rows(rows)

    def validate_config(self, config: dict) -> list:
        """Validate required fields."""
        errors = []
        for field in self.REQUIRED_FIELDS:
            value = config.get(field)
            if value is None or (isinstance(value, str) and value.strip() == ""):
                errors.append(f"缺少必填配置项: {field}")

        environment = config.get("Environment", "").strip().lower()
        if environment not in {"test", "prod"}:
            errors.append("Environment 必须为 test 或 prod")

        if environment == "prod" and not config.get("Prod_API_Base_URL", "").strip():
            errors.append("Environment=prod 时必须填写 Prod_API_Base_URL")

        return errors

    def _ensure_config_dir(self) -> None:
        """Ensure the config directory exists."""
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir)

    def _read_config_rows(self) -> list:
        """Read the config file as raw 3-column rows."""
        existing_df = pd.read_csv(self.config_path, names=["A", "B", "C"])
        return [
            [
                self._normalize_cell(row["A"]),
                self._normalize_cell(row["B"]),
                self._normalize_cell(row["C"]),
            ]
            for _, row in existing_df.iterrows()
        ]

    def _build_synced_rows(self, existing_rows: list) -> list:
        """Build the synced config rows, preserving existing values."""
        existing_map = {row[0]: row for row in existing_rows}
        merged_rows = []
        for key, default_value, comment in self.DEFAULT_TEMPLATE:
            if key in existing_map and key != "基础信息":
                merged_rows.append([key, existing_map[key][1], comment])
            else:
                merged_rows.append([key, default_value, comment])
        return merged_rows

    def _write_rows(self, rows: list) -> None:
        """Write rows to the config file."""
        self._ensure_config_dir()
        pd.DataFrame(rows).to_csv(
            self.config_path,
            index=False,
            header=False,
            encoding="utf_8_sig",
        )

    @staticmethod
    def _normalize_cell(value):
        """Normalize a CSV cell to a string."""
        if isinstance(value, float) and math.isnan(value):
            return ""
        return str(value)
