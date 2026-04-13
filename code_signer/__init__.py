# -*- coding: utf-8 -*-
"""
代码签名模块
使用方法：
    from code_signer import CodeSigner

    signer = CodeSigner.from_config('config.py')
    success, message = signer.sign_file('app.exe')

    # 或使用便捷函数
    from code_signer import sign_file
    success, message = sign_file('app.exe', 'certificate_name')
"""

from .core import CodeSigner
from .config import SigningConfig, CertificateConfig, ToolConfig
from .utils import find_signing_tools, verify_signature

# 便捷函数
def sign_file(file_path: str, config_path: str = None, certificate_name: str = None):
    """便捷的文件签名函数"""
    signer = CodeSigner.from_config(config_path) if config_path else CodeSigner()
    return signer.sign_file(file_path, certificate_name)

def verify_file_signature(file_path: str):
    """便捷的签名验证函数"""
    return verify_signature(file_path)

__version__ = "1.0.0"
__all__ = [
    'CodeSigner',
    'SigningConfig',
    'CertificateConfig',
    'ToolConfig',
    'find_signing_tools',
    'verify_signature',
    'sign_file',
    'verify_file_signature'
]