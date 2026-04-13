# Code Signer - 现代化代码签名模块

[![Python Version](https://img.shields.io/badge/python-3.7+-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

一个现代化的代码签名模块，提供智能配置加载、多工具支持和类型安全的配置管理。设计用于企业级应用，同时保持向后兼容性。

## ✨ 主要特性

- 🧠 **智能配置加载** - 支持多种配置格式和优先级
- 🔒 **类型安全配置** - Python配置文件，支持IDE自动补全和验证
- 🛠️ **多工具支持** - signtool、PowerShell、osslsigncode
- 🔄 **向后兼容** - 与现有JSON配置完全兼容
- 📝 **配置验证** - 自动验证配置完整性
- 🚀 **独立运行** - 可作为命令行工具使用
- 📦 **即插即用** - 简单的API接口
- 📊 **详细记录** - 完整的签名历史和日志

## 🏗️ 架构设计

### 智能配置加载系统

```
配置加载优先级:
1. 指定配置路径
2. code_signer/examples/project_config.py
3. signature_config.json (兼容模式)
4. 内置默认配置
```

### 模块结构

```
code_signer/
├── __init__.py              # 模块入口和主要API
├── config.py               # 配置类定义 (SigningConfig等)
├── config_loader.py        # 智能配置加载器
├── core.py                 # 签名核心逻辑
├── utils.py                # 工具函数 (safe_subprocess_run等)
├── cli.py                  # 命令行接口
├── examples/               # 配置示例
│   ├── project_config.py   # 项目配置模板
│   └── README.md           # 配置示例说明
└── README.md              # 本文档
```

## 🚀 快速开始

### 安装和集成

将 `code_signer` 文件夹复制到您的项目中：

```bash
cp -r code_signer /path/to/your/project/
```

### 基本使用

#### 方式1：智能配置加载 (推荐)

```python
from code_signer.config_loader import load_signing_config
from code_signer import CodeSigner

# 智能加载最佳配置
config = load_signing_config()
signer = CodeSigner(config)

# 签名文件
success, message = signer.sign_file("my_app.exe")
print(f"签名结果: {success}, {message}")
```

#### 方式2：指定配置文件

```python
from code_signer import CodeSigner

# 使用指定的Python配置文件
signer = CodeSigner.from_config("my_config.py")
success, message = signer.sign_file("my_app.exe")
```

#### 方式3：传统方式 (兼容)

```python
from signing_tool import SigningTool

# 自动加载配置 (向后兼容)
tool = SigningTool()
success, message = tool.sign_file("my_app.exe")
```

## 📝 配置系统

### Python配置文件 (推荐)

创建配置文件 `my_config.py`：

```python
# -*- coding: utf-8 -*-
from code_signer.config import SigningConfig, CertificateConfig, ToolConfig
from code_signer.config import FilePathsConfig, PoliciesConfig, OutputConfig

# 创建配置实例
CONFIG = SigningConfig()

# 基本配置
CONFIG.enabled = True
CONFIG.default_certificate = "my_app_cert"
CONFIG.timestamp_server = "http://timestamp.digicert.com"
CONFIG.hash_algorithm = "sha256"

# 证书配置
app_cert = CertificateConfig(
    name="my_app_cert",
    sha1="144ac4069565211ab67d25a9d6d33af0e18e511e",
    subject="CN=My Application",
    issuer="CN=Certificate Authority",
    valid_from="2025-01-01",
    valid_to="2027-01-01",
    description="我的应用程序证书"
)

CONFIG.add_certificate(app_cert)

# 策略配置
CONFIG.policies = PoliciesConfig(
    verify_before_sign=True,
    backup_before_sign=False,
    auto_retry=True,
    max_retries=3,
    record_signing_history=True
)

# 文件路径配置
CONFIG.file_paths = FilePathsConfig(
    search_patterns=["dist/*.exe", "*.exe"],
    exclude_patterns=["*.tmp.exe", "*_unsigned.exe"],
    record_directory="./signature_records"
)
```

### JSON配置文件 (兼容)

```json
{
  "signature": {
    "enabled": true,
    "default_certificate": "my_app_cert",
    "timestamp_server": "http://timestamp.digicert.com",
    "hash_algorithm": "sha256",
    "certificates": {
      "my_app_cert": {
        "name": "My App Certificate",
        "sha1": "144ac4069565211ab67d25a9d6d33af0e18e511e",
        "subject": "CN=My Application",
        "issuer": "CN=Certificate Authority"
      }
    }
  }
}
```

## 🔧 API参考

### CodeSigner 类

主要的签名器类，提供完整的签名功能。

#### 构造函数

```python
CodeSigner(config: Optional[SigningConfig] = None)
```

**参数**:
- `config`: 签名配置对象，如果为None则使用智能配置加载

#### 主要方法

##### sign_file()

```python
sign_file(file_path: str, certificate_name: Optional[str] = None) -> Tuple[bool, str]
```

**参数**:
- `file_path`: 要签名的文件路径
- `certificate_name`: 证书名称，如果为None使用默认证书

**返回**:
- `Tuple[bool, str]`: (是否成功, 消息)

##### sign_files()

```python
sign_files(file_paths: List[str], certificate_name: Optional[str] = None) -> List[Tuple[str, bool, str]]
```

**参数**:
- `file_paths`: 要签名的文件路径列表
- `certificate_name`: 证书名称

**返回**:
- `List[Tuple[str, bool, str]]`: [(文件路径, 是否成功, 消息)]

##### verify_signature()

```python
verify_signature(file_path: str) -> Tuple[bool, str]
```

**参数**:
- `file_path`: 要验证的文件路径

**返回**:
- `Tuple[bool, str]`: (是否验证成功, 消息)

### 配置加载器

#### load_signing_config()

```python
load_signing_config(config_path: Optional[str] = None) -> SigningConfig
```

智能加载签名配置，支持多种格式和优先级。

**参数**:
- `config_path`: 指定的配置文件路径

**返回**:
- `SigningConfig`: 配置对象

#### get_config_load_info()

```python
get_config_load_info() -> Dict[str, Any]
```

获取配置加载信息。

**返回**:
- `Dict[str, Any]`: 包含加载源、配置类型等信息的字典

### 配置类

#### SigningConfig

主配置类，包含所有签名相关配置。

**主要属性**:
- `enabled: bool` - 是否启用签名功能
- `default_certificate: str` - 默认证书名称
- `timestamp_server: str` - 时间戳服务器
- `hash_algorithm: str` - 哈希算法 (sha1/sha256)
- `certificates: Dict[str, CertificateConfig]` - 证书配置
- `signing_tools: Dict[str, ToolConfig]` - 签名工具配置
- `policies: PoliciesConfig` - 策略配置
- `file_paths: FilePathsConfig` - 文件路径配置
- `output: OutputConfig` - 输出配置

#### CertificateConfig

证书配置类。

**必需参数**:
- `name: str` - 证书名称
- `sha1: str` - 证书SHA1指纹

**可选参数**:
- `subject: str` - 证书主题
- `issuer: str` - 证书颁发者
- `valid_from: str` - 有效期开始
- `valid_to: str` - 有效期结束
- `description: str` - 证书描述

## 🛠️ 签名工具支持

### 1. SignTool (Windows SDK)

**优先级**: 1 (最高)

**特点**:
- Microsoft官方工具
- 最佳兼容性
- 支持时间戳

**自动查找路径**:
- Windows Kits 10/bin/
- Windows Kits 8.1/bin/
- Visual Studio安装目录

### 2. PowerShell

**优先级**: 2

**特点**:
- Windows内置
- 无需额外安装
- 使用Set-AuthenticodeSignature

### 3. osslsigncode

**优先级**: 3

**特点**:
- 跨平台支持
- 开源工具
- 需要单独安装

## 📊 配置验证

### 自动验证

配置加载时会自动验证：

- ✅ 证书SHA1不为空
- ✅ 默认证书存在于证书列表
- ✅ 哈希算法有效性
- ✅ 最大重试次数大于0
- ✅ 工具名称不为空

### 手动验证

```python
from code_signer.config_loader import load_signing_config

config = load_signing_config()
errors = config.validate()

if errors:
    print("配置验证失败:")
    for error in errors:
        print(f"  - {error}")
else:
    print("配置验证通过")
```

## 🔄 迁移指南

### 从 signing_tool.py 迁移

#### 旧代码

```python
from signing_tool import SigningTool

tool = SigningTool("signature_config.json")
success, message = tool.sign_file("app.exe")
```

#### 新代码 (推荐)

```python
from code_signer.config_loader import load_signing_config
from code_signer import CodeSigner

config = load_signing_config()
signer = CodeSigner(config)
success, message = signer.sign_file("app.exe")
```

### 从JSON配置迁移到Python配置

#### 步骤1: 备份现有配置

```bash
cp signature_config.json signature_config.json.backup
```

#### 步骤2: 创建Python配置

```python
# my_config.py
from code_signer.config import SigningConfig, CertificateConfig

CONFIG = SigningConfig()
CONFIG.enabled = True

# 从JSON迁移证书配置
cert = CertificateConfig(
    name="my_cert",
    sha1="your_sha1_here",
    # ... 其他配置
)
CONFIG.add_certificate(cert)
```

#### 步骤3: 验证新配置

```python
from code_signer.config_loader import load_signing_config
config = load_signing_config("my_config.py")
errors = config.validate()
```

## 🐛 故障排除

### 常见问题

#### 1. 配置加载失败

**错误**: `配置文件不存在`

**解决方案**:
```python
from code_signer.config_loader import load_signing_config, get_config_load_info

config = load_signing_config()
info = get_config_load_info()
print(f"使用的配置源: {info['load_source']}")
```

#### 2. 证书验证失败

**错误**: `证书SHA1不能为空`

**解决方案**:
```python
# 确保证书配置正确
cert = CertificateConfig(
    name="my_cert",
    sha1="actual_sha1_value",  # 必须提供真实值
    subject="CN=My App"
)
```

#### 3. 签名工具未找到

**错误**: `未找到可用的签名工具`

**解决方案**:
- 安装Windows 10 SDK (用于signtool)
- 系统会自动降级到PowerShell
- 或手动指定工具路径

#### 4. 编码问题

**错误**: `UnicodeDecodeError`

**解决方案**:
```python
import os
os.environ['DEBUG_ENCODING'] = '1'  # 启用编码调试
```

### 调试工具

#### 配置测试

```bash
python test_signing_config.py
```

#### 签名功能测试

```bash
python test_actual_signing.py
```

## 📚 最佳实践

### 1. 配置管理
- 使用Python配置文件获得类型安全
- 定期验证配置完整性
- 使用版本控制管理配置变更
- 备份重要配置文件

### 2. 证书管理
- 定期检查证书有效期
- 使用描述性的证书名称
- 安全保管证书信息
- 在证书到期前更新配置

### 3. 错误处理
- 检查所有操作的返回值
- 记录详细的错误日志
- 实现适当的重试机制
- 提供用户友好的错误信息

### 4. 安全考虑
- 保护证书私钥安全
- 验证签名工具的完整性
- 使用可信的时间戳服务器
- 定期审计签名记录

## 🤝 贡献指南

欢迎提交问题和改进建议！

### 开发环境设置

```bash
# 克隆项目
git clone <repository_url>
cd project

# 安装依赖
pip install -r requirements.txt

# 运行测试
python test_signing_config.py
python test_actual_signing.py
```

### 代码规范

- 遵循PEP 8代码风格
- 添加类型注解
- 编写单元测试
- 更新文档

## 📄 许可证

MIT License - 详见 [LICENSE](LICENSE) 文件

## 🔗 相关链接

- [配置文件详细说明](../04-配置文件说明.md)
- [签名使用说明](../06-签名使用说明.md)
- [项目总览](../01-项目总览和使用指南.md)
- [迁移指南](MIGRATION.md)