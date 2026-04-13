# -*- coding: utf-8 -*-
"""
默认签名配置示例
这是一个通用的配置模板，适用于大多数应用程序
"""
from code_signer.config import SigningConfig, CertificateConfig

# 创建配置实例
CONFIG = SigningConfig()

# 启用签名功能
CONFIG.enabled = True

# 默认证书名称
CONFIG.default_certificate = "my_app"

# 时间戳服务器
CONFIG.timestamp_server = "http://timestamp.digicert.com"

# 哈希算法 (sha1, sha256)
CONFIG.hash_algorithm = "sha256"

# 添加证书配置
cert = CertificateConfig(
    name="my_app",
    sha1="YOUR_CERTIFICATE_SHA1_HERE",  # 替换为您的证书SHA1
    subject="CN=My Application",         # 可选：证书使用者
    issuer="CN=Certificate Authority",   # 可选：证书颁发者
    description="我的应用程序证书"        # 可选：证书描述
)
CONFIG.add_certificate(cert)

# 输出配置
CONFIG.output.verbose = True  # 显示详细输出
CONFIG.output.save_records = True  # 保存签名记录