# 迁移指南 - 从 signing_tool.py 到 code_signer

本指南将帮助您从现有的 `signing_tool.py` 迁移到新的模块化 `code_signer`。

## 🎯 迁移目标

- ✅ **保持功能完整性** - 所有现有功能都得到保留
- ✅ **向后兼容** - 现有代码无需立即修改
- ✅ **渐进式迁移** - 可以逐步迁移到新的配置方式
- ✅ **增强配置能力** - 利用 Python 配置的优势

## 📋 迁移检查清单

### 阶段 1：准备迁移
- [ ] 备份现有的签名配置
- [ ] 检查当前使用的签名工具和证书
- [ ] 确认现有的签名流程

### 阶段 2：配置迁移
- [ ] 转换 JSON 配置为 Python 配置
- [ ] 验证新配置的正确性
- [ ] 测试配置加载

### 阶段 3：代码迁移
- [ ] 更新导入语句
- [ ] 适配新的 API
- [ ] 测试签名功能

### 阶段 4：验证部署
- [ ] 功能测试
- [ ] 性能验证
- [ ] 文档更新

## 🔄 配置转换

### 现有 JSON 配置 (signature_config.json)

```json
{
  "signature": {
    "enabled": true,
    "default_certificate": "pdf_rename_operation",
    "timestamp_server": "http://timestamp.digicert.com",
    "hash_algorithm": "sha256",
    "certificates": {
      "pdf_rename_operation": {
        "name": "PDF_Rename_Operation证书",
        "sha1": "144ac4069565211ab67d25a9d6d33af0e18e511e",
        "subject": "CN=PDF_Rename_Operation...",
        "issuer": "CN=TUVSUD-IssuingCA...",
        "description": "TUVSUD颁发的PDF重命名工具专用证书"
      }
    }
  }
}
```

### 转换后的 Python 配置 (signing_config.py)

```python
from code_signer.config import SigningConfig, CertificateConfig

# 创建配置实例
CONFIG = SigningConfig()

# 基本配置
CONFIG.enabled = True
CONFIG.default_certificate = "pdf_rename_operation"
CONFIG.timestamp_server = "http://timestamp.digicert.com"
CONFIG.hash_algorithm = "sha256"

# 证书配置
cert = CertificateConfig(
    name="pdf_rename_operation",
    sha1="144ac4069565211ab67d25a9d6d33af0e18e511e",
    subject="CN=PDF_Rename_Operation, OU=PS:Softlines, O=TÜV SÜD Certification and Testing (China) Co. Ltd. Xiamen Branch, L=Xiamen, C=CN",
    issuer="CN=TUVSUD-IssuingCA, O=TUVSUD, C=SG",
    description="TUVSUD颁发的PDF重命名工具专用证书"
)
CONFIG.add_certificate(cert)
```

### 自动转换脚本

```python
# convert_config.py - 自动转换脚本
import json
from pathlib import Path

def convert_json_to_python(json_path: str, output_path: str):
    """将 JSON 配置转换为 Python 配置"""

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    signature_config = data.get('signature', {})

    # 生成 Python 配置代码
    config_code = '''# -*- coding: utf-8 -*-
"""
自动生成的签名配置文件
由 signing_tool.py 配置转换而来
"""
from code_signer.config import SigningConfig, CertificateConfig, ToolConfig
from code_signer.config import FilePathsConfig, PoliciesConfig, OutputConfig

# 创建配置实例
CONFIG = SigningConfig()

# 基本配置
CONFIG.enabled = {enabled}
CONFIG.default_certificate = "{default_cert}"
CONFIG.timestamp_server = "{timestamp_server}"
CONFIG.hash_algorithm = "{hash_algorithm}"
'''

    # 转换证书配置
    certificates = signature_config.get('certificates', {})
    cert_code = []

    for name, cert_data in certificates.items():
        cert_code.append(f'''
cert_{name} = CertificateConfig(
    name="{name}",
    sha1="{cert_data.get('sha1', '')}",
    subject="{cert_data.get('subject', '')}",
    issuer="{cert_data.get('issuer', '')}",
    valid_from="{cert_data.get('valid_from', '')}",
    valid_to="{cert_data.get('valid_to', '')}",
    description="{cert_data.get('description', '')}"
)
CONFIG.add_certificate(cert_{name})
''')

    config_code += '\n'.join(cert_code)

    # 保存 Python 配置文件
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(config_code.format(
            enabled=signature_config.get('enabled', True),
            default_cert=signature_config.get('default_certificate', 'default'),
            timestamp_server=signature_config.get('timestamp_server', 'http://timestamp.digicert.com'),
            hash_algorithm=signature_config.get('hash_algorithm', 'sha256')
        ))

    print(f"配置已转换: {json_path} -> {output_path}")

# 使用示例
if __name__ == "__main__":
    convert_json_to_python("signature_config.json", "signing_config.py")
```

## 💻 代码迁移

### 1. 导入语句变更

**之前:**
```python
from signing_tool import SigningTool

tool = SigningTool("signature_config.json")
success, message = tool.sign_file("app.exe", "certificate_name")
```

**之后:**
```python
from code_signer import CodeSigner

signer = CodeSigner.from_config("signing_config.py")
success, message = signer.sign_file("app.exe", "certificate_name")
```

### 2. 便捷函数使用

**之前:**
```python
from signing_tool import sign_file as sign_file_legacy

success, message = sign_file_legacy("app.exe")
```

**之后:**
```python
from code_signer import sign_file

success, message = sign_file("app.py", "signing_config.py", "certificate_name")
```

### 3. 批量操作

**之前:**
```python
from signing_tool import SigningTool

tool = SigningTool("signature_config.json")
results = tool.batch_sign("./dist/")
```

**之后:**
```python
from code_signer import CodeSigner

signer = CodeSigner.from_config("signing_config.py")
results = signer.batch_sign("./dist/", "certificate_name")
```

## 🔧 渐进式迁移策略

### 阶段 1: 并行运行 (推荐)

保持现有系统不变，并行测试新模块：

```python
# 现有代码保持不变
from signing_tool import SigningTool as LegacySigner

# 新模块测试
try:
    from code_signer import CodeSigner
    NEW_SIGNER_AVAILABLE = True
except ImportError:
    NEW_SIGNER_AVAILABLE = False

def sign_file_fallback(file_path: str, certificate_name: str = None):
    """带回退的签名函数"""
    if NEW_SIGNER_AVAILABLE:
        try:
            signer = CodeSigner.from_config("signing_config.py")
            return signer.sign_file(file_path, certificate_name)
        except Exception as e:
            print(f"[警告] 新模块签名失败，回退到旧模块: {e}")

    # 回退到旧模块
    tool = LegacySigner("signature_config.json")
    return tool.sign_file(file_path, certificate_name)
```

### 阶段 2: 配置文件转换

1. 运行转换脚本生成 Python 配置
2. 验证配置加载是否正常
3. 测试签名功能

```python
# 验证配置
def validate_new_config():
    try:
        from code_signer import load_config_from_file
        config = load_config_from_file("signing_config.py")
        print("✅ 新配置验证成功")
        return True
    except Exception as e:
        print(f"❌ 新配置验证失败: {e}")
        return False
```

### 阶段 3: 完全迁移

```python
# 完全迁移到新模块
from code_signer import CodeSigner

def main():
    signer = CodeSigner.from_config("signing_config.py")
    success, message = signer.sign_file("app.exe")

    if success:
        print(f"签名成功: {message}")
    else:
        print(f"签名失败: {message}")
```

## 🧪 测试迁移

### 功能测试脚本

```python
# test_migration.py
import os
import sys
from pathlib import Path

def test_legacy_signer():
    """测试旧签名工具"""
    try:
        from signing_tool import SigningTool
        tool = SigningTool("signature_config.json")
        print("✅ 旧模块加载成功")

        # 测试配置
        cert_config = tool.get_certificate_config("pdf_rename_operation")
        if cert_config:
            print("✅ 旧模块证书配置正常")
        else:
            print("❌ 旧模块证书配置失败")

        return True
    except Exception as e:
        print(f"❌ 旧模块测试失败: {e}")
        return False

def test_new_signer():
    """测试新签名工具"""
    try:
        from code_signer import CodeSigner
        signer = CodeSigner.from_config("code_signer/examples/project_config.py")
        print("✅ 新模块加载成功")

        # 测试配置
        cert_config = signer.config.get_certificate("pdf_rename_operation")
        if cert_config:
            print("✅ 新模块证书配置正常")
        else:
            print("❌ 新模块证书配置失败")

        return True
    except Exception as e:
        print(f"❌ 新模块测试失败: {e}")
        return False

def test_file_signing():
    """测试文件签名功能"""
    test_file = "test_file.exe"

    # 创建测试文件
    with open(test_file, "wb") as f:
        f.write(b"test content")

    try:
        # 测试新模块
        from code_signer import CodeSigner
        signer = CodeSigner.from_config("code_signer/examples/project_config.py")

        # 注意：这里不会实际签名，只测试配置和初始化
        print("✅ 新模块签名测试准备完成")

    except Exception as e:
        print(f"❌ 签名测试失败: {e}")

    finally:
        # 清理测试文件
        if os.path.exists(test_file):
            os.remove(test_file)

def main():
    print("=== 迁移测试开始 ===")

    # 测试旧模块
    print("\n1. 测试旧模块:")
    legacy_ok = test_legacy_signer()

    # 测试新模块
    print("\n2. 测试新模块:")
    new_ok = test_new_signer()

    # 测试签名功能
    print("\n3. 测试签名功能:")
    test_file_signing()

    print("\n=== 测试总结 ===")
    print(f"旧模块: {'✅ 正常' if legacy_ok else '❌ 异常'}")
    print(f"新模块: {'✅ 正常' if new_ok else '❌ 异常'}")

    if new_ok:
        print("\n🎉 可以安全迁移到新模块!")
    else:
        print("\n⚠️  请先解决新模块的问题再迁移")

if __name__ == "__main__":
    main()
```

## 🚨 常见问题和解决方案

### Q1: 配置文件加载失败

**问题:** `FileNotFoundError: 配置文件不存在`

**解决方案:**
```python
# 检查配置文件路径
import os
config_paths = [
    "signing_config.py",
    "code_signer/examples/project_config.py",
    "./signing_config.py"
]

for path in config_paths:
    if os.path.exists(path):
        print(f"找到配置文件: {path}")
        break
```

### Q2: 证书配置错误

**问题:** `ValueError: 证书SHA1不能为空`

**解决方案:**
```python
# 验证证书配置
from code_signer import load_config_from_file

config = load_config_from_file("signing_config.py")
cert = config.get_certificate("your_cert_name")
if cert and cert.sha1:
    print("证书配置正常")
else:
    print("请检查证书SHA1配置")
```

### Q3: 签名工具找不到

**问题:** `[错误] 未找到可用的签名工具`

**解决方案:**
```bash
# 检查可用工具
python -m code_signer.cli tools

# 或者手动检查
where signtool
```

### Q4: 向后兼容性问题

**问题:** 现有代码无法运行

**解决方案:**
```python
# 使用兼容性包装器
class SignerAdapter:
    def __init__(self, config_path=None):
        try:
            from code_signer import CodeSigner
            self.signer = CodeSigner.from_config(config_path)
            self.use_new = True
        except ImportError:
            from signing_tool import SigningTool
            self.signer = SigningTool(config_path)
            self.use_new = False

    def sign_file(self, file_path, cert_name=None):
        if self.use_new:
            return self.signer.sign_file(file_path, cert_name)
        else:
            return self.signer.sign_file(file_path, cert_name)
```

## ✅ 迁移验证清单

### 功能验证
- [ ] 证书配置正确加载
- [ ] 签名工具正常工作
- [ ] 文件签名功能正常
- [ ] 签名验证功能正常
- [ ] 批量签名功能正常
- [ ] 错误处理正确

### 性能验证
- [ ] 签名速度无显著下降
- [ ] 内存使用正常
- [ ] 启动时间可接受

### 兼容性验证
- [ ] 现有脚本无需修改即可运行
- [ ] 配置文件转换正确
- [ ] API 接口兼容

### 部署验证
- [ ] 打包工具集成正常
- [ ] CI/CD 流程正常
- [ ] 生产环境运行正常

---

## 🎉 迁移完成

恭喜！您已成功迁移到新的模块化签名系统。现在您可以享受：

- 🔧 更灵活的配置管理
- 🚀 更好的模块化设计
- 📦 更易于集成和维护
- 🛠️ 更强大的功能特性

如有任何问题，请查看 [README.md](README.md) 或提交 Issue。