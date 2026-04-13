# -*- coding: utf-8 -*-
"""
Python配置文件管理器
支持类型安全的配置定义和加载
使用方法：
    config = SigningConfig()
    config.add_certificate(CertificateConfig(
        name='my_cert',
        sha1='abc123...',
        subject='CN=My App',
        issuer='CN=CA'
    ))
"""
import os
import sys
from typing import Dict, List, Optional, Any, Union
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class CertificateConfig:
    """证书配置"""
    name: str
    sha1: str
    subject: str = ""
    issuer: str = ""
    valid_from: str = ""
    valid_to: str = ""
    description: str = ""

    def __post_init__(self):
        if not self.name:
            raise ValueError("证书名称不能为空")
        if not self.sha1:
            raise ValueError("证书SHA1不能为空")


@dataclass
class ToolConfig:
    """签名工具配置"""
    name: str
    enabled: bool = True
    path: str = "auto"  # "auto" 或具体路径
    priority: int = 999
    description: str = ""

    def __post_init__(self):
        if not self.name:
            raise ValueError("工具名称不能为空")


@dataclass
class FilePathsConfig:
    """文件路径配置"""
    search_patterns: List[str] = field(default_factory=lambda: ["*.exe"])
    exclude_patterns: List[str] = field(default_factory=lambda: ["*.tmp.exe", "*_unsigned.exe"])
    record_directory: str = "./signature_records"


@dataclass
class PoliciesConfig:
    """策略配置"""
    verify_before_sign: bool = True
    backup_before_sign: bool = False
    auto_retry: bool = True
    max_retries: int = 3
    record_signing_history: bool = True


@dataclass
class OutputConfig:
    """输出配置"""
    verbose: bool = True
    save_records: bool = True
    record_format: str = "json"
    create_log_file: bool = True


@dataclass
class SigningConfig:
    """签名配置主类"""
    enabled: bool = True
    default_certificate: str = "default"
    timestamp_server: str = "http://timestamp.digicert.com"
    hash_algorithm: str = "sha256"

    # 子配置
    certificates: Dict[str, CertificateConfig] = field(default_factory=dict)
    signing_tools: Dict[str, ToolConfig] = field(default_factory=dict)
    file_paths: FilePathsConfig = field(default_factory=FilePathsConfig)
    policies: PoliciesConfig = field(default_factory=PoliciesConfig)
    output: OutputConfig = field(default_factory=OutputConfig)

    def __post_init__(self):
        """初始化默认配置"""
        if not self.signing_tools:
            self._add_default_tools()

    def _add_default_tools(self):
        """添加默认签名工具配置"""
        default_tools = [
            ToolConfig(
                name="signtool",
                enabled=True,
                path="auto",
                priority=1,
                description="Windows SDK signtool.exe"
            ),
            ToolConfig(
                name="powershell",
                enabled=True,
                priority=2,
                description="PowerShell Set-AuthenticodeSignature"
            ),
            ToolConfig(
                name="osslsigncode",
                enabled=True,
                path="auto",
                priority=3,
                description="osslsigncode工具"
            )
        ]

        for tool in default_tools:
            self.signing_tools[tool.name] = tool

    def add_certificate(self, cert: CertificateConfig):
        """添加证书配置"""
        self.certificates[cert.name] = cert

    def get_certificate(self, name: str) -> Optional[CertificateConfig]:
        """获取证书配置"""
        return self.certificates.get(name)

    def add_tool(self, tool: ToolConfig):
        """添加签名工具配置"""
        self.signing_tools[tool.name] = tool

    def get_tool(self, name: str) -> Optional[ToolConfig]:
        """获取工具配置"""
        return self.signing_tools.get(name)

    def get_enabled_tools(self) -> List[ToolConfig]:
        """获取启用的工具，按优先级排序"""
        enabled_tools = [tool for tool in self.signing_tools.values() if tool.enabled]
        return sorted(enabled_tools, key=lambda x: x.priority)

    def validate(self) -> List[str]:
        """验证配置，返回错误信息列表"""
        errors = []

        if self.enabled and not self.certificates:
            errors.append("签名功能已启用但未配置任何证书")

        if self.default_certificate not in self.certificates:
            errors.append(f"默认证书 '{self.default_certificate}' 未在证书列表中找到")

        if self.hash_algorithm not in ["sha1", "sha256"]:
            errors.append(f"不支持的哈希算法: {self.hash_algorithm}")

        if self.policies.max_retries < 1:
            errors.append("最大重试次数必须大于0")

        return errors

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'SigningConfig':
        """从字典创建配置对象"""
        config = cls()

        # 更新基本配置
        for key in ['enabled', 'default_certificate', 'timestamp_server', 'hash_algorithm']:
            if key in data:
                setattr(config, key, data[key])

        # 加载证书配置
        if 'certificates' in data:
            for name, cert_data in data['certificates'].items():
                # 移除 name 避免重复参数
                cert_data_copy = cert_data.copy()
                cert_data_copy.pop('name', None)
                cert = CertificateConfig(name=name, **cert_data_copy)
                config.add_certificate(cert)

        # 加载工具配置
        if 'signing_tools' in data:
            for name, tool_data in data['signing_tools'].items():
                # 移除 name 避免重复参数
                tool_data_copy = tool_data.copy()
                tool_data_copy.pop('name', None)
                tool = ToolConfig(name=name, **tool_data_copy)
                config.add_tool(tool)

        # 加载文件路径配置
        if 'file_paths' in data:
            config.file_paths = FilePathsConfig(**data['file_paths'])

        # 加载策略配置
        if 'policies' in data:
            config.policies = PoliciesConfig(**data['policies'])

        # 加载输出配置
        if 'output' in data:
            config.output = OutputConfig(**data['output'])

        return config

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            'enabled': self.enabled,
            'default_certificate': self.default_certificate,
            'timestamp_server': self.timestamp_server,
            'hash_algorithm': self.hash_algorithm,
            'certificates': {name: {
                'sha1': cert.sha1,
                'subject': cert.subject,
                'issuer': cert.issuer,
                'valid_from': cert.valid_from,
                'valid_to': cert.valid_to,
                'description': cert.description
            } for name, cert in self.certificates.items()},
            'signing_tools': {name: {
                'enabled': tool.enabled,
                'path': tool.path,
                'priority': tool.priority,
                'description': tool.description
            } for name, tool in self.signing_tools.items()},
            'file_paths': {
                'search_patterns': self.file_paths.search_patterns,
                'exclude_patterns': self.file_paths.exclude_patterns,
                'record_directory': self.file_paths.record_directory
            },
            'policies': {
                'verify_before_sign': self.policies.verify_before_sign,
                'backup_before_sign': self.policies.backup_before_sign,
                'auto_retry': self.policies.auto_retry,
                'max_retries': self.policies.max_retries,
                'record_signing_history': self.policies.record_signing_history
            },
            'output': {
                'verbose': self.output.verbose,
                'save_records': self.output.save_records,
                'record_format': self.output.record_format,
                'create_log_file': self.output.create_log_file
            }
        }


def load_config_from_file(config_path: str) -> SigningConfig:
    """
    从Python文件加载配置
    配置文件应该包含一个名为 CONFIG 的 SigningConfig 实例
    """
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件不存在: {config_path}")

    # 动态导入配置文件
    config_dir = os.path.dirname(os.path.abspath(config_path))
    config_filename = os.path.basename(config_path).replace('.py', '')

    # 将配置文件目录添加到Python路径
    if config_dir not in sys.path:
        sys.path.insert(0, config_dir)

    try:
        config_module = __import__(config_filename)
        if not hasattr(config_module, 'CONFIG'):
            raise AttributeError(f"配置文件 {config_path} 必须包含 CONFIG 变量")

        config = getattr(config_module, 'CONFIG')
        if not isinstance(config, SigningConfig):
            raise TypeError("CONFIG 必须是 SigningConfig 实例")

        # 验证配置
        errors = config.validate()
        if errors:
            raise ValueError(f"配置验证失败:\n" + "\n".join(f"  - {error}" for error in errors))

        return config

    finally:
        # 清理Python路径
        if config_dir in sys.path:
            sys.path.remove(config_dir)


def load_config_from_module(module_path: str) -> SigningConfig:
    """
    从模块加载配置
    module_path 格式: "package.module"
    """
    try:
        module = __import__(module_path, fromlist=['CONFIG'])
        config = getattr(module, 'CONFIG')

        if not isinstance(config, SigningConfig):
            raise TypeError("CONFIG 必须是 SigningConfig 实例")

        errors = config.validate()
        if errors:
            raise ValueError(f"配置验证失败:\n" + "\n".join(f"  - {error}" for error in errors))

        return config

    except ImportError as e:
        raise ImportError(f"无法导入配置模块 {module_path}: {e}")
    except AttributeError as e:
        raise AttributeError(f"模块 {module_path} 必须包含 CONFIG 变量: {e}")