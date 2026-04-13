# -*- coding: utf-8 -*-
"""
统一的配置管理模块
提供配置加载、验证和管理功能
"""

import os
import sys
import json
import logging
from typing import Dict, Any, Optional
from dataclasses import dataclass, asdict

@dataclass
class UpdateSettings:
    """更新设置数据类"""
    github_repo: str = "chen-huai/Temu_PDF_Rename_APP"
    github_api_base: str = "https://api.github.com"
    update_check_interval_days: int = 30
    auto_check_enabled: bool = True
    backup_count: int = 3
    download_timeout: int = 300
    max_retries: int = 3
    enable_prerelease: bool = False
    current_version: str = "1.0.0"

class SettingsManager:
    """设置管理器"""

    def __init__(self, config_file: Optional[str] = None):
        """
        初始化设置管理器
        :param config_file: 配置文件路径
        """
        self.logger = logging.getLogger(__name__)

        # 配置文件路径
        if config_file is None:
            # 获取可执行文件所在目录
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))

            config_file = os.path.join(base_dir, "update_config.json")

        self.config_file = config_file
        self.settings = UpdateSettings()

        # 加载配置
        self._load_settings()

    def _load_settings(self):
        """加载配置"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)

                # 更新设置
                for key, value in config_data.items():
                    if hasattr(self.settings, key):
                        setattr(self.settings, key, value)

                self.logger.info("配置加载成功")
            else:
                self.logger.info("配置文件不存在，使用默认设置")
                self._save_settings()  # 创建默认配置文件

        except Exception as e:
            self.logger.error(f"加载配置失败: {e}")
            self.logger.info("使用默认配置")

    def _save_settings(self):
        """保存配置"""
        try:
            # 确保目录存在
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)

            # 转换为字典并保存
            config_dict = asdict(self.settings)

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_dict, f, indent=2, ensure_ascii=False)

            self.logger.info("配置保存成功")

        except Exception as e:
            self.logger.error(f"保存配置失败: {e}")

    def get_settings(self) -> UpdateSettings:
        """获取设置"""
        return self.settings

    def update_setting(self, key: str, value: Any):
        """
        更新单个设置
        :param key: 设置键
        :param value: 设置值
        """
        if hasattr(self.settings, key):
            setattr(self.settings, key, value)
            self._save_settings()
        else:
            self.logger.warning(f"未知的设置键: {key}")

    def validate_settings(self) -> tuple:
        """
        验证设置
        :return: (是否有效, 错误信息)
        """
        errors = []

        # 验证GitHub仓库格式
        if not self.settings.github_repo or '/' not in self.settings.github_repo:
            errors.append("GitHub仓库格式无效")

        # 验证API基础URL
        if not self.settings.github_api_base.startswith('https://'):
            errors.append("GitHub API基础URL格式无效")

        # 验证间隔时间
        if self.settings.update_check_interval_days < 1:
            errors.append("更新检查间隔必须大于0天")

        # 验证超时时间
        if self.settings.download_timeout < 10:
            errors.append("下载超时时间不能少于10秒")

        # 验证重试次数
        if self.settings.max_retries < 0 or self.settings.max_retries > 10:
            errors.append("重试次数必须在0-10之间")

        # 验证备份数量
        if self.settings.backup_count < 1 or self.settings.backup_count > 10:
            errors.append("备份数量必须在1-10之间")

        # 验证版本格式
        if not self._is_valid_version(self.settings.current_version):
            errors.append("当前版本格式无效")

        is_valid = len(errors) == 0
        error_message = "; ".join(errors) if errors else ""

        return is_valid, error_message

    def _is_valid_version(self, version: str) -> bool:
        """
        验证版本号格式
        :param version: 版本号
        :return: 是否有效
        """
        try:
            from packaging import version as ver
            ver.parse(version.lstrip('v'))
            return True
        except:
            return False

    def get_github_urls(self) -> Dict[str, str]:
        """
        获取GitHub相关URL
        :return: URL字典
        """
        base_url = f"{self.settings.github_api_base}/repos/{self.settings.github_repo}"
        return {
            'releases_url': f"{base_url}/releases",
            'latest_release_url': f"{base_url}/releases/latest",
            'repo_url': f"https://github.com/{self.settings.github_repo}",
            'download_base_url': f"https://github.com/{self.settings.github_repo}/releases/download"
        }

    def reset_to_defaults(self):
        """重置为默认设置"""
        self.settings = UpdateSettings()
        self._save_settings()
        self.logger.info("配置已重置为默认值")

    def export_settings(self) -> str:
        """
        导出设置为JSON字符串
        :return: JSON字符串
        """
        return json.dumps(asdict(self.settings), indent=2, ensure_ascii=False)

    def import_settings(self, json_string: str) -> tuple:
        """
        从JSON字符串导入设置
        :param json_string: JSON字符串
        :return: (是否成功, 错误信息)
        """
        try:
            config_data = json.loads(json_string)

            # 验证并应用设置
            temp_settings = UpdateSettings()
            for key, value in config_data.items():
                if hasattr(temp_settings, key):
                    setattr(temp_settings, key, value)

            # 验证新设置
            self.settings = temp_settings
            is_valid, error_msg = self.validate_settings()

            if is_valid:
                self._save_settings()
                return True, "设置导入成功"
            else:
                # 恢复原设置
                self._load_settings()
                return False, f"设置验证失败: {error_msg}"

        except json.JSONDecodeError as e:
            return False, f"JSON格式错误: {str(e)}"
        except Exception as e:
            return False, f"导入设置失败: {str(e)}"

# 全局设置管理器实例
_settings_manager = None

def get_settings_manager() -> SettingsManager:
    """获取全局设置管理器实例"""
    global _settings_manager
    if _settings_manager is None:
        _settings_manager = SettingsManager()
    return _settings_manager

def get_settings() -> UpdateSettings:
    """获取当前设置"""
    return get_settings_manager().get_settings()