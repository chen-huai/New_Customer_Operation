# -*- coding: utf-8 -*-
"""
命令行接口模块
提供独立的命令行签名工具
使用方法：
    python -m code_signer.cli sign app.exe
    python -m code_signer.cli --config config.py --cert my_cert batch ./dist/
"""
import argparse
import sys
import os
from typing import List, Optional

from .core import CodeSigner
from .config import SigningConfig, CertificateConfig, load_config_from_file
from .utils import find_signing_tools, get_system_info, format_file_size


def create_parser() -> argparse.ArgumentParser:
    """创建命令行参数解析器"""
    parser = argparse.ArgumentParser(
        prog='code_signer',
        description='代码签名工具 - 支持多种签名工具和配置方式',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  # 签名单个文件
  python -m code_signer.cli sign app.exe

  # 使用指定配置文件签名
  python -m code_signer.cli --config myconfig.py sign app.exe

  # 使用指定证书签名
  python -m code_signer.cli --cert my_cert sign app.exe

  # 批量签名目录中的文件
  python -m code_signer.cli batch ./dist/

  # 验证文件签名
  python -m code_signer.cli verify app.exe

  # 显示证书信息
  python -m code_signer.cli cert-info --cert my_cert

  # 查看可用工具
  python -m code_signer.cli tools

  # 生成配置模板
  python -m code_signer.cli init-config --output myconfig.py
        """
    )

    # 全局参数
    parser.add_argument('--config', '-c',
                       help='配置文件路径')
    parser.add_argument('--cert',
                       help='证书名称')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='详细输出')
    parser.add_argument('--version', action='version', version='%(prog)s 1.0.0')

    # 子命令
    subparsers = parser.add_subparsers(dest='command', help='可用命令')

    # sign 命令
    sign_parser = subparsers.add_parser('sign', help='签名单个文件')
    sign_parser.add_argument('file', help='要签名的文件路径')

    # batch 命令
    batch_parser = subparsers.add_parser('batch', help='批量签名文件')
    batch_parser.add_argument('directory', nargs='?', default='.',
                             help='要签名的目录路径（默认为当前目录）')

    # verify 命令
    verify_parser = subparsers.add_parser('verify', help='验证文件签名')
    verify_parser.add_argument('file', help='要验证的文件路径')

    # cert-info 命令
    cert_info_parser = subparsers.add_parser('cert-info', help='显示证书信息')
    cert_info_parser.add_argument('--cert',
                                 help='证书名称（默认显示默认证书）')

    # tools 命令
    subparsers.add_parser('tools', help='显示可用签名工具')

    # init-config 命令
    init_parser = subparsers.add_parser('init-config', help='生成配置模板')
    init_parser.add_argument('--output', '-o', default='signing_config.py',
                            help='输出配置文件路径（默认：signing_config.py）')
    init_parser.add_argument('--type', choices=['basic', 'advanced'],
                            default='basic', help='配置类型（默认：basic）')

    # info 命令
    subparsers.add_parser('info', help='显示系统信息')

    return parser


def cmd_sign(args) -> int:
    """签名单个文件命令"""
    try:
        # 创建签名器
        if args.config:
            signer = CodeSigner.from_config(args.config)
        else:
            signer = CodeSigner()

        # 设置详细输出
        if args.verbose:
            signer.config.output.verbose = True

        # 检查文件
        if not os.path.exists(args.file):
            print(f"[错误] 文件不存在: {args.file}")
            return 1

        file_size = os.path.getsize(args.file)
        print(f"开始签名文件: {args.file}")
        print(f"文件大小: {format_file_size(file_size)}")

        # 执行签名
        success, message = signer.sign_file(args.file, args.cert)

        if success:
            print(f"[成功] {message}")
            return 0
        else:
            print(f"[失败] {message}")
            return 1

    except Exception as e:
        print(f"[错误] 签名过程中发生异常: {e}")
        return 1


def cmd_batch(args) -> int:
    """批量签名命令"""
    try:
        # 创建签名器
        if args.config:
            signer = CodeSigner.from_config(args.config)
        else:
            signer = CodeSigner()

        # 设置详细输出
        if args.verbose:
            signer.config.output.verbose = True

        # 查找目标文件
        print(f"搜索目录: {args.directory}")
        files = signer.find_target_files(args.directory)

        if not files:
            print("[信息] 未找到要签名的文件")
            return 0

        print(f"找到 {len(files)} 个文件:")
        for file_path in files:
            file_size = os.path.getsize(file_path)
            print(f"  - {file_path} ({format_file_size(file_size)})")

        # 确认操作
        if not args.verbose:
            choice = input("\n是否继续签名? (y/n): ").strip().lower()
            if choice not in ['y', 'yes', '是']:
                print("操作已取消")
                return 0

        # 执行批量签名
        print("\n开始批量签名...")
        results = signer.batch_sign(args.directory, args.cert)

        # 统计结果
        success_count = sum(1 for success, _ in results.values() if success)
        total_count = len(results)

        print(f"\n签名完成! 成功: {success_count}/{total_count}")

        # 显示失败的文件
        failed_files = [(file_path, message) for file_path, (success, message) in results.items() if not success]
        if failed_files:
            print("\n失败的文件:")
            for file_path, message in failed_files:
                print(f"  - {file_path}: {message}")

        return 0 if success_count == total_count else 1

    except Exception as e:
        print(f"[错误] 批量签名过程中发生异常: {e}")
        return 1


def cmd_verify(args) -> int:
    """验证签名命令"""
    try:
        from .utils import verify_signature

        if not os.path.exists(args.file):
            print(f"[错误] 文件不存在: {args.file}")
            return 1

        print(f"验证文件签名: {args.file}")
        success, message = verify_signature(args.file)

        if success:
            print("[成功] 文件签名有效")
            if message.strip():
                print("签名详情:")
                for line in message.strip().split('\n'):
                    print(f"  {line}")
            return 0
        else:
            print(f"[失败] {message}")
            return 1

    except Exception as e:
        print(f"[错误] 验证过程中发生异常: {e}")
        return 1


def cmd_cert_info(args) -> int:
    """显示证书信息命令"""
    try:
        # 创建签名器
        if args.config:
            signer = CodeSigner.from_config(args.config)
        else:
            signer = CodeSigner()

        # 显示证书信息
        signer.display_certificate_info(args.cert)
        return 0

    except Exception as e:
        print(f"[错误] 获取证书信息失败: {e}")
        return 1


def cmd_tools(args) -> int:
    """显示可用工具命令"""
    try:
        print("可用的签名工具:")
        tools = find_signing_tools()

        if not tools:
            print("  未找到可用的签名工具")
            return 1

        for name, path in tools.items():
            print(f"  {name}: {path}")

        return 0

    except Exception as e:
        print(f"[错误] 获取工具信息失败: {e}")
        return 1


def cmd_init_config(args) -> int:
    """生成配置模板命令"""
    try:
        config_content = generate_config_template(args.type)

        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(config_content)

        print(f"配置模板已生成: {args.output}")
        print(f"配置类型: {args.type}")
        print("请根据需要修改配置文件中的参数")
        return 0

    except Exception as e:
        print(f"[错误] 生成配置模板失败: {e}")
        return 1


def cmd_info(args) -> int:
    """显示系统信息命令"""
    try:
        info = get_system_info()
        print("系统信息:")
        for key, value in info.items():
            print(f"  {key}: {value}")

        print("\n可用签名工具:")
        tools = find_signing_tools()
        for name, path in tools.items():
            print(f"  {name}: {path}")

        return 0

    except Exception as e:
        print(f"[错误] 获取系统信息失败: {e}")
        return 1


def generate_config_template(config_type: str) -> str:
    """生成配置模板内容"""
    if config_type == 'basic':
        return '''# -*- coding: utf-8 -*-
"""
基础签名配置文件
请根据您的实际情况修改以下配置
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
'''
    else:  # advanced
        return '''# -*- coding: utf-8 -*-
"""
高级签名配置文件
包含完整的配置选项
"""
from code_signer.config import SigningConfig, CertificateConfig, ToolConfig
from code_signer.config import FilePathsConfig, PoliciesConfig, OutputConfig

# 创建配置实例
CONFIG = SigningConfig()

# 基本配置
CONFIG.enabled = True
CONFIG.default_certificate = "production_app"
CONFIG.timestamp_server = "http://timestamp.digicert.com"
CONFIG.hash_algorithm = "sha256"

# 证书配置
prod_cert = CertificateConfig(
    name="production_app",
    sha1="PRODUCTION_CERT_SHA1_HERE",
    subject="CN=My Production App",
    issuer="CN=Production CA",
    valid_from="2024-01-01",
    valid_to="2026-01-01",
    description="生产环境应用程序证书"
)

dev_cert = CertificateConfig(
    name="development_app",
    sha1="DEVELOPMENT_CERT_SHA1_HERE",
    subject="CN=My Development App",
    issuer="CN=Development CA",
    description="开发环境应用程序证书"
)

CONFIG.add_certificate(prod_cert)
CONFIG.add_certificate(dev_cert)

# 签名工具配置
# Windows SDK signtool (优先级1)
signtool_config = ToolConfig(
    name="signtool",
    enabled=True,
    path="auto",  # 自动查找或指定路径如 "C:\\path\\to\\signtool.exe"
    priority=1,
    description="Windows SDK signtool.exe"
)

# PowerShell (优先级2)
powershell_config = ToolConfig(
    name="powershell",
    enabled=True,
    priority=2,
    description="PowerShell Set-AuthenticodeSignature"
)

# osslsigncode (优先级3)
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

# 文件路径配置
CONFIG.file_paths = FilePathsConfig(
    search_patterns=["*.exe", "*.dll", "*.cab"],  # 搜索的文件类型
    exclude_patterns=["*.tmp.exe", "*_unsigned.exe"],  # 排除的文件类型
    record_directory="./signature_records"  # 签名记录保存目录
)

# 策略配置
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
'''


def main(argv: List[str] = None) -> int:
    """主函数"""
    parser = create_parser()
    args = parser.parse_args(argv)

    # 如果没有指定命令，显示帮助
    if not args.command:
        parser.print_help()
        return 1

    # 路由到对应的命令处理函数
    command_handlers = {
        'sign': cmd_sign,
        'batch': cmd_batch,
        'verify': cmd_verify,
        'cert-info': cmd_cert_info,
        'tools': cmd_tools,
        'init-config': cmd_init_config,
        'info': cmd_info
    }

    handler = command_handlers.get(args.command)
    if handler:
        return handler(args)
    else:
        print(f"[错误] 未知命令: {args.command}")
        return 1


if __name__ == '__main__':
    sys.exit(main())