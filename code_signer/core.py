# -*- coding: utf-8 -*-
"""
核心签名逻辑模块
基于原有的 signing_tool.py 重构，使用新的配置系统
使用方法：
    from code_signer import CodeSigner, SigningConfig

    config = SigningConfig()
    signer = CodeSigner(config)
    success, message = signer.sign_file('app.exe')
"""
import os
import sys
import subprocess
import time
import json
import glob
import hashlib
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any

from .config import SigningConfig, CertificateConfig, ToolConfig
from .utils import find_signtool, find_osslsigncode, calculate_file_hash, safe_subprocess_run


class SigningRecord:
    """签名记录类"""

    def __init__(self, file_path: str, certificate_name: str, success: bool,
                 message: str, tool_name: str, certificate_sha1: str = ""):
        self.file_path = file_path
        self.file_hash = calculate_file_hash(file_path)
        self.certificate_name = certificate_name
        self.certificate_sha1 = certificate_sha1
        self.signing_time = time.strftime("%Y-%m-%d %H:%M:%S")
        self.success = success
        self.message = message
        self.tool = tool_name

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            "file_path": self.file_path,
            "file_hash": self.file_hash,
            "certificate_name": self.certificate_name,
            "certificate_sha1": self.certificate_sha1,
            "signing_time": self.signing_time,
            "success": self.success,
            "message": self.message,
            "tool": self.tool
        }


class CodeSigner:
    """代码签名器主类"""

    def __init__(self, config: SigningConfig = None):
        """
        初始化签名器
        :param config: 签名配置，如果为None则使用默认配置
        """
        self.config = config or SigningConfig()
        self.signing_records: List[SigningRecord] = []

        # 验证配置
        if not self.config.enabled:
            raise ValueError("签名功能未启用，请检查配置中的 enabled 设置")

        errors = self.config.validate()
        if errors:
            raise ValueError(f"配置验证失败:\n" + "\n".join(f"  - {error}" for error in errors))

    @classmethod
    def from_config(cls, config_path: str = None) -> 'CodeSigner':
        """
        从配置文件创建签名器
        支持Python (.py) 和 JSON (.json) 配置文件
        :param config_path: 配置文件路径，如果为None则使用默认配置
        :return: CodeSigner实例
        """
        if config_path:
            # 使用 config_loader 智能加载配置
            from .config_loader import load_signing_config
            config = load_signing_config(config_path)
            return cls(config)
        else:
            return cls()

    @classmethod
    def from_module(cls, module_path: str) -> 'CodeSigner':
        """
        从模块创建签名器
        :param module_path: 模块路径，如 "myapp.signing_config"
        :return: CodeSigner实例
        """
        from .config import load_config_from_module
        config = load_config_from_module(module_path)
        return cls(config)

    def find_available_tool(self, tool_name: str = None) -> Optional[str]:
        """
        查找可用的签名工具
        :param tool_name: 指定工具名称，如果为None则按优先级查找第一个可用的
        :return: 工具路径或名称，如果未找到返回None
        """
        if tool_name:
            tool_config = self.config.get_tool(tool_name)
            if not tool_config or not tool_config.enabled:
                return None

            if tool_name == 'signtool':
                return find_signtool(tool_config.path)
            elif tool_name == 'powershell':
                return 'powershell'
            elif tool_name == 'osslsigncode':
                return find_osslsigncode(tool_config.path)
            else:
                return None
        else:
            # 按优先级查找第一个可用的工具
            for tool_config in self.config.get_enabled_tools():
                tool_path = self.find_available_tool(tool_config.name)
                if tool_path:
                    return tool_path
            return None

    def verify_certificate_exists(self, cert_config: CertificateConfig) -> bool:
        """
        验证证书是否存在
        :param cert_config: 证书配置
        :return: 是否存在
        """
        if not cert_config or not cert_config.sha1:
            return False

        try:
            cmd = ['certutil', '-user', '-store', 'My', cert_config.sha1]
            result = safe_subprocess_run(cmd, encoding='utf-8')
            return result.returncode == 0
        except Exception:
            return False

    def sign_with_signtool(self, file_path: str, cert_config: CertificateConfig) -> Tuple[bool, str]:
        """使用signtool签名文件"""
        signtool_path = find_signtool('auto')
        if not signtool_path:
            return False, "[错误] 未找到signtool.exe"

        cmd = [
            signtool_path, "sign",
            "/sha1", cert_config.sha1,
            "/fd", self.config.hash_algorithm,
            "/td", self.config.hash_algorithm,
            "/tr", self.config.timestamp_server,
            file_path
        ]

        try:
            if self.config.output.verbose:
                print(f"[执行] {' '.join(cmd)}")

            result = safe_subprocess_run(cmd, encoding='utf-8')

            if result.returncode == 0:
                return True, "[成功] signtool签名成功"
            else:
                return False, f"[错误] signtool签名失败: {result.stderr}"

        except Exception as e:
            return False, f"[错误] signtool签名异常: {e}"

    def sign_with_powershell(self, file_path: str, cert_config: CertificateConfig) -> Tuple[bool, str]:
        """使用PowerShell签名文件"""
        script = f'''
        Set-AuthenticodeSignature -FilePath "{file_path}" -Certificate "{cert_config.sha1}" -TimestampServer "{self.config.timestamp_server}"
        $LASTEXITCODE
        '''

        try:
            result = safe_subprocess_run(['powershell', '-Command', script], encoding='utf-8')

            if result.returncode == 0:
                return True, "[成功] PowerShell签名成功"
            else:
                return False, f"[错误] PowerShell签名失败: {result.stderr}"

        except Exception as e:
            return False, f"[错误] PowerShell签名异常: {e}"

    def sign_with_osslsigncode(self, file_path: str, cert_config: CertificateConfig) -> Tuple[bool, str]:
        """使用osslsigncode签名文件"""
        osslsigncode_path = find_osslsigncode('auto')
        if not osslsigncode_path:
            return False, "[错误] 未找到osslsigncode"

        output_path = file_path + ".signed"
        cmd = [
            osslsigncode_path, "sign",
            "-certs", cert_config.sha1,
            "-t", self.config.timestamp_server,
            "-h", self.config.hash_algorithm,
            "-in", file_path,
            "-out", output_path
        ]

        try:
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0 and os.path.exists(output_path):
                # 替换原文件
                os.replace(output_path, file_path)
                return True, "[成功] osslsigncode签名成功"
            else:
                return False, f"[错误] osslsigncode签名失败: {result.stderr}"

        except Exception as e:
            return False, f"[错误] osslsigncode签名异常: {e}"

    def sign_file(self, file_path: str, certificate_name: str = None) -> Tuple[bool, str]:
        """
        签名文件
        :param file_path: 文件路径
        :param certificate_name: 证书名称，如果为None则使用默认证书
        :return: (是否成功, 消息)
        """
        if not os.path.exists(file_path):
            return False, f"[错误] 文件不存在: {file_path}"

        # 获取证书配置
        if not certificate_name:
            certificate_name = self.config.default_certificate

        cert_config = self.config.get_certificate(certificate_name)
        if not cert_config:
            return False, f"[错误] 未找到证书配置: {certificate_name}"

        # 验证证书是否存在
        if self.config.policies.verify_before_sign:
            if not self.verify_certificate_exists(cert_config):
                return False, f"[错误] 证书不存在或无法访问: {cert_config.name}"

        # 验证文件是否已签名
        if self.config.policies.verify_before_sign:
            is_signed, _ = self.verify_signature(file_path)
            if is_signed:
                return False, f"[警告] 文件已有签名: {file_path}"

        # 获取签名工具
        tool_path = self.find_available_tool()
        if not tool_path:
            return False, "[错误] 未找到可用的签名工具"

        # 执行签名
        success = False
        message = ""
        tool_name = ""

        max_retries = self.config.policies.max_retries
        auto_retry = self.config.policies.auto_retry

        for attempt in range(max_retries):
            if tool_path.endswith('signtool.exe'):
                tool_name = "signtool"
                success, message = self.sign_with_signtool(file_path, cert_config)
            elif tool_path == 'powershell':
                tool_name = "powershell"
                success, message = self.sign_with_powershell(file_path, cert_config)
            elif 'osslsigncode' in tool_path:
                tool_name = "osslsigncode"
                success, message = self.sign_with_osslsigncode(file_path, cert_config)
            else:
                return False, f"[错误] 不支持的签名工具: {tool_path}"

            if success or not auto_retry:
                break

            if attempt < max_retries - 1:
                print(f"[重试] 签名失败，{attempt + 1}秒后重试...")
                time.sleep(attempt + 1)

        # 保存签名记录
        if success and self.config.policies.record_signing_history:
            record = SigningRecord(
                file_path=file_path,
                certificate_name=certificate_name,
                success=success,
                message=message,
                tool_name=tool_name,
                certificate_sha1=cert_config.sha1
            )
            self.save_signing_record(record)

        return success, message

    def verify_signature(self, file_path: str) -> Tuple[bool, str]:
        """
        验证文件签名
        :param file_path: 文件路径
        :return: (是否成功, 消息)
        """
        signtool_path = find_signtool('auto')
        if not signtool_path:
            return False, "[错误] 未找到signtool.exe"

        try:
            cmd = [signtool_path, "verify", "/pa", file_path]
            result = safe_subprocess_run(cmd, encoding='utf-8')

            if result.returncode == 0:
                return True, result.stdout
            else:
                return False, result.stderr

        except Exception as e:
            return False, f"[错误] 验证异常: {e}"

    def save_signing_record(self, record: SigningRecord):
        """
        保存签名记录
        :param record: 签名记录
        """
        self.signing_records.append(record)

        if not self.config.output.save_records:
            return

        record_dir = self.config.file_paths.record_directory
        os.makedirs(record_dir, exist_ok=True)

        record_file = os.path.join(record_dir,
            f"{os.path.basename(record.file_path)}_signing_record.json")

        try:
            with open(record_file, "w", encoding="utf-8") as f:
                json.dump(record.to_dict(), f, indent=2, ensure_ascii=False)

            if self.config.output.verbose:
                print(f"[记录] 签名记录已保存到: {record_file}")
        except Exception as e:
            print(f"[警告] 保存签名记录失败: {e}")

    def find_target_files(self, search_dir: str = ".") -> List[str]:
        """
        查找目标文件
        :param search_dir: 搜索目录
        :return: 文件路径列表
        """
        files = []
        search_patterns = self.config.file_paths.search_patterns
        exclude_patterns = self.config.file_paths.exclude_patterns

        for pattern in search_patterns:
            pattern_path = os.path.join(search_dir, pattern)
            for file_path in glob.glob(pattern_path):
                # 检查排除模式
                if any(glob.fnmatch.fnmatch(os.path.basename(file_path), exclude_pattern)
                      for exclude_pattern in exclude_patterns):
                    continue
                files.append(file_path)

        return files

    def batch_sign(self, search_dir: str = ".", certificate_name: str = None) -> Dict[str, Tuple[bool, str]]:
        """
        批量签名文件
        :param search_dir: 搜索目录
        :param certificate_name: 证书名称
        :return: 文件路径到结果的映射
        """
        files = self.find_target_files(search_dir)
        if not files:
            return {}

        results = {}
        for file_path in files:
            print(f"\n[处理] 签名文件: {file_path}")
            success, message = self.sign_file(file_path, certificate_name)
            results[file_path] = (success, message)

            if success:
                print(f"[成功] {message}")
            else:
                print(f"[失败] {message}")

        return results

    def display_certificate_info(self, certificate_name: str = None):
        """
        显示证书信息
        :param certificate_name: 证书名称，如果为None则显示默认证书
        """
        if not certificate_name:
            certificate_name = self.config.default_certificate

        cert_config = self.config.get_certificate(certificate_name)
        if not cert_config:
            print(f"[错误] 未找到证书配置: {certificate_name}")
            return

        print("\n" + "="*60)
        print(f"证书信息: {cert_config.name}")
        print("="*60)
        print(f"证书名称: {cert_config.name}")
        print(f"SHA1哈希: {cert_config.sha1}")
        print(f"使用者: {cert_config.subject}")
        print(f"颁发者: {cert_config.issuer}")
        if cert_config.valid_from and cert_config.valid_to:
            print(f"有效期: {cert_config.valid_from} - {cert_config.valid_to}")
        print(f"描述: {cert_config.description}")
        print("="*60)