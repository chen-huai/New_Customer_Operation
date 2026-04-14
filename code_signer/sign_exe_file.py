# -*- coding: utf-8 -*-
"""
EXE 数字签名工具
================
独立签名模块，通过 SHA1 直接定位证书，无需外部配置文件。
放置于 code_signer/ 包内，随包一起复制到其他项目即可复用。

使用方法（作为包导入）：
    from code_signer.sign_exe_file import sign_exe_with_sha1, verify_exe_signature
    success, message = sign_exe_with_sha1("dist/MyApp.exe")

使用方法（直接运行）：
    python code_signer/sign_exe_file.py dist/MyApp.exe
"""

import os
import sys
from pathlib import Path

# 兼容"包导入"和"直接运行"两种方式
try:
    from .core import CodeSigner
    from .config import SigningConfig, CertificateConfig, PoliciesConfig
except ImportError:
    sys.path.insert(0, str(Path(__file__).parent.parent))
    from code_signer.core import CodeSigner
    from code_signer.config import SigningConfig, CertificateConfig, PoliciesConfig


# ====================== 证书配置（内嵌，无需外部文件）======================
# 复制到其他项目时，修改此处的 SHA1 即可
CERT_SHA1        = "144ac4069565211ab67d25a9d6d33af0e18e511e"
CERT_NAME        = "pdf_rename_operation"
TIMESTAMP_SERVER = "http://timestamp.digicert.com"
HASH_ALGORITHM   = "sha256"

# 直接运行时的默认目标文件
DEFAULT_EXE      = "dist/Sap_Operate_theme.exe"


def _build_signer() -> CodeSigner:
    """构建 CodeSigner 实例，证书配置完全内嵌"""
    config = SigningConfig()
    config.default_certificate = CERT_NAME
    config.timestamp_server    = TIMESTAMP_SERVER
    config.hash_algorithm      = HASH_ALGORITHM
    config.policies = PoliciesConfig(
        verify_before_sign=False,   # 不预检，避免重复运行时误报"已签名"
        backup_before_sign=False,
        auto_retry=True,
        max_retries=3,
        record_signing_history=True,
    )
    config.add_certificate(CertificateConfig(
        name=CERT_NAME,
        sha1=CERT_SHA1,
        description="代码签名证书（TUVSUD颁发）",
    ))
    return CodeSigner(config)


def sign_exe_with_sha1(exe_path: str) -> tuple:
    """
    使用 SHA1 证书对 EXE 文件进行数字签名。

    :param exe_path: 目标 EXE 文件路径
    :return: (是否成功, 消息)
    """
    if not os.path.exists(exe_path):
        return False, f"文件不存在: {exe_path}"

    print(f"\n[签名] 目标文件: {exe_path}")
    print(f"[签名] 证书 SHA1: {CERT_SHA1}")

    try:
        success, message = _build_signer().sign_file(exe_path, CERT_NAME)
        status = "[OK] 成功" if success else "[FAIL] 失败"
        print(f"[签名] {status}: {message}")
        return success, message
    except Exception as e:
        print(f"[签名] [ERROR] 异常: {e}")
        return False, str(e)


def verify_exe_signature(exe_path: str) -> tuple:
    """
    验证 EXE 文件的数字签名。

    :param exe_path: 目标 EXE 文件路径
    :return: (是否有效, 消息)
    """
    if not os.path.exists(exe_path):
        return False, f"文件不存在: {exe_path}"

    try:
        is_signed, message = _build_signer().verify_signature(exe_path)
        status = "[OK] 签名有效" if is_signed else f"[WARN] {message}"
        print(f"[验证] {status}")
        return is_signed, message
    except Exception as e:
        print(f"[验证] [ERROR] 异常: {e}")
        return False, str(e)


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_EXE

    if not os.path.exists(target):
        print(f"[ERROR] 文件不存在: {target}")
        print("用法: python code_signer/sign_exe_file.py <exe_path>")
        sys.exit(1)

    ok, _ = sign_exe_with_sha1(target)
    if ok:
        verify_exe_signature(target)
    sys.exit(0 if ok else 1)
