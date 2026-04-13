# -*- coding: utf-8 -*-
"""
PDF重命名工具项目专用配置示例
基于原项目 signature_config.json 的配置转换
"""
from code_signer.config import SigningConfig, CertificateConfig, ToolConfig
from code_signer.config import FilePathsConfig, PoliciesConfig, OutputConfig

# 创建配置实例
CONFIG = SigningConfig()

# 基本配置
CONFIG.enabled = True
CONFIG.default_certificate = "pdf_rename_operation"
CONFIG.timestamp_server = "http://timestamp.digicert.com"
CONFIG.hash_algorithm = "sha256"

# 证书配置 - PDF重命名工具专用证书
pdf_cert = CertificateConfig(
    name="pdf_rename_operation",
    sha1="144ac4069565211ab67d25a9d6d33af0e18e511e",
    subject="CN=PDF_Rename_Operation, OU=PS:Softlines, O=TÜV SÜD Certification and Testing (China) Co. Ltd. Xiamen Branch, L=Xiamen, C=CN",
    issuer="CN=TUVSUD-IssuingCA, O=TUVSUD, C=SG",
    valid_from="2025-10-15",
    valid_to="2027-10-15",
    description="TUVSUD颁发的PDF重命名工具专用证书"
)

# 默认证书模板 - 使用实际可用的证书信息
default_cert = CertificateConfig(
    name="default_template",
    sha1="144ac4069565211ab67d25a9d6d33af0e18e511e",  # 使用与pdf_cert相同的证书
    subject="CN=PDF_Rename_Operation, OU=PS:Softlines, O=TÜV SÜD Certification and Testing (China) Co. Ltd. Xiamen Branch, L=Xiamen, C=CN",
    issuer="CN=TUVSUD-IssuingCA, O=TUVSUD, C=SG",
    description="默认证书配置，基于TUVSUD颁发的证书"
)

CONFIG.add_certificate(pdf_cert)
CONFIG.add_certificate(default_cert)

# 签名工具配置 - 保持与原配置一致的优先级
signtool_config = ToolConfig(
    name="signtool",
    enabled=True,
    path="auto",  # 自动查找
    priority=1,
    description="Windows SDK signtool.exe"
)

powershell_config = ToolConfig(
    name="powershell",
    enabled=True,
    priority=2,
    description="PowerShell Set-AuthenticodeSignature"
)

ossl_config = ToolConfig(
    name="osslsigncode",
    enabled=True,
    path="auto",
    priority=3,
    description="osslsigncode工具"
)

CONFIG.add_tool(signtool_config)
CONFIG.add_tool(powershell_config)
CONFIG.add_tool(ossl_config)

# 文件路径配置 - 针对PDF重命名工具项目
CONFIG.file_paths = FilePathsConfig(
    search_patterns=[
        "dist/*.exe",      # PyInstaller输出目录
        "*.exe",           # 当前目录的exe文件
        "build/*.exe",     # 构建目录
        "PDF重命名工具_便携版/*.exe"  # 便携版目录
    ],
    exclude_patterns=[
        "*.tmp.exe",       # 临时文件
        "*_unsigned.exe",  # 未签名文件
        "*_backup*.exe"    # 备份文件
    ],
    record_directory="./signature_records"
)

# 策略配置 - 保持与原配置一致
CONFIG.policies = PoliciesConfig(
    verify_before_sign=True,      # 签名前验证文件是否已签名
    backup_before_sign=False,     # 签名前备份文件
    auto_retry=True,              # 失败时自动重试
    max_retries=3,                # 最大重试次数
    record_signing_history=True   # 记录签名历史
)

# 输出配置
CONFIG.output = OutputConfig(
    verbose=True,              # 显示详细输出
    save_records=True,         # 保存签名记录
    record_format="json",      # 记录格式
    create_log_file=True       # 创建日志文件
)

# 配置验证
def validate_config():
    """验证配置是否正确"""
    errors = CONFIG.validate()
    if errors:
        print("[配置错误] 发现以下问题:")
        for error in errors:
            print(f"  - {error}")
        return False
    else:
        print("[配置成功] 所有配置项验证通过")
        return True

# 自动验证配置
if __name__ == "__main__":
    validate_config()