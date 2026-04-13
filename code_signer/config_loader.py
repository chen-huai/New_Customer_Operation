# -*- coding: utf-8 -*-
"""
配置加载器
智能配置加载机制，支持多种配置格式
"""
import os
import sys
import json
import importlib.util
from typing import Optional, Any, Dict, Union
from pathlib import Path

from .config import SigningConfig


class ConfigLoader:
    """智能配置加载器"""

    def __init__(self):
        self.loaded_config: Optional[SigningConfig] = None
        self.load_source: str = ""

    def load_config(
        self,
        config_path: Optional[str] = None,
        search_paths: Optional[list] = None
    ) -> SigningConfig:
        """
        加载配置文件

        优先级：
        1. 指定的配置文件路径
        2. code_signer/examples/project_config.py
        3. signature_config.json (兼容模式)
        4. 默认配置

        :param config_path: 指定的配置文件路径
        :param search_paths: 搜索路径列表
        :return: SigningConfig实例
        """
        if search_paths is None:
            search_paths = [
                "code_signer/examples/project_config.py",
                "signature_config.json",
            ]

        # 如果指定了配置路径，优先加载
        if config_path:
            config = self._load_specific_config(config_path)
            if config:
                self.loaded_config = config
                self.load_source = config_path
                return config

        # 按优先级尝试加载预设配置文件
        for path in search_paths:
            config = self._load_specific_config(path)
            if config:
                self.loaded_config = config
                self.load_source = path
                return config

        # 所有配置都加载失败，使用默认配置
        print("[警告] 所有配置文件加载失败，使用默认配置")
        default_config = SigningConfig()
        self.loaded_config = default_config
        self.load_source = "默认配置"
        return default_config

    def _load_specific_config(self, config_path: str) -> Optional[SigningConfig]:
        """
        加载指定的配置文件

        :param config_path: 配置文件路径
        :return: SigningConfig实例或None
        """
        if not os.path.exists(config_path):
            return None

        try:
            if config_path.endswith('.py'):
                return self._load_python_config(config_path)
            elif config_path.endswith('.json'):
                return self._load_json_config(config_path)
            else:
                # 根据文件扩展名自动判断
                if config_path.endswith('.py') or 'project_config' in config_path:
                    return self._load_python_config(config_path)
                else:
                    return self._load_json_config(config_path)
        except Exception as e:
            print(f"[警告] 加载配置文件 {config_path} 失败: {e}")
            return None

    def _load_python_config(self, config_path: str) -> Optional[SigningConfig]:
        """
        加载Python配置文件

        :param config_path: Python配置文件路径
        :return: SigningConfig实例或None
        """
        try:
            # 动态导入Python模块
            spec = importlib.util.spec_from_file_location("config_module", config_path)
            if spec is None or spec.loader is None:
                return None

            config_module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(config_module)

            # 查找CONFIG变量
            if hasattr(config_module, 'CONFIG'):
                config = getattr(config_module, 'CONFIG')
                if isinstance(config, SigningConfig):
                    print(f"[信息] 使用Python配置: {config_path}")
                    return config

            # 如果没有CONFIG变量，尝试创建默认配置
            print(f"[警告] Python配置文件 {config_path} 中没有找到CONFIG变量")
            return None

        except Exception as e:
            print(f"[错误] 加载Python配置失败: {e}")
            return None

    def _load_json_config(self, config_path: str) -> Optional[SigningConfig]:
        """
        加载JSON配置文件（兼容模式）

        :param config_path: JSON配置文件路径
        :return: SigningConfig实例或None
        """
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            # 转换JSON配置为SigningConfig
            config = SigningConfig.from_dict(data)
            print(f"[信息] 使用JSON配置: {config_path}")
            return config

        except Exception as e:
            print(f"[错误] 加载JSON配置失败: {e}")
            return None

    def get_load_info(self) -> Dict[str, Any]:
        """
        获取配置加载信息

        :return: 包含加载信息的字典
        """
        return {
            'load_source': self.load_source,
            'config_loaded': self.loaded_config is not None,
            'config_type': type(self.loaded_config).__name__ if self.loaded_config else None
        }


# 全局配置加载器实例
_config_loader = ConfigLoader()


def load_signing_config(config_path: Optional[str] = None) -> SigningConfig:
    """
    加载签名配置

    :param config_path: 配置文件路径
    :return: SigningConfig实例
    """
    global _config_loader
    return _config_loader.load_config(config_path)


def get_config_load_info() -> Dict[str, Any]:
    """
    获取配置加载信息

    :return: 配置加载信息
    """
    global _config_loader
    return _config_loader.get_load_info()


# 使用示例
if __name__ == "__main__":
    # 测试配置加载
    config = load_signing_config()
    load_info = get_config_load_info()

    print(f"配置加载完成:")
    print(f"  加载源: {load_info['load_source']}")
    print(f"  配置类型: {load_info['config_type']}")

    # 验证配置
    errors = config.validate()
    if errors:
        print("配置验证失败:")
        for error in errors:
            print(f"  - {error}")
    else:
        print("配置验证成功")