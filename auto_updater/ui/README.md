# Auto Updater UI Module

## 概述

这是一个完全解耦的自动更新功能UI模块，提供独立的更新界面组件，可与任何PyQt5应用程序集成。

## 特性

- ✅ **完全解耦**：UI与业务逻辑完全分离
- ✅ **易于集成**：一行代码即可集成到现有项目
- ✅ **高度可定制**：支持文本、样式和行为的自定义
- ✅ **向后兼容**：保持与现有代码的完全兼容性
- ✅ **错误处理**：完善的异常处理和回退机制
- ✅ **资源管理**：统一的文本和样式资源管理

## 快速开始

### 基本使用

```python
from PyQt5.QtWidgets import QMainWindow, QApplication
from auto_updater import AutoUpdater

# 创建主窗口
app = QApplication([])
main_window = QMainWindow()

# 一行代码完成更新功能集成
updater = AutoUpdater(main_window)
success = updater.setup_update_ui(main_window.menuBar())

if success:
    print("更新功能集成成功！")
    main_window.show()
    app.exec_()
else:
    print("更新功能集成失败")
```

### 高级使用

```python
from auto_updater import AutoUpdater
from auto_updater.ui import UpdateUIManager, UpdateUIText

# 自定义应用信息
UpdateUIText.APP_NAME = "我的应用程序"
UpdateUIText.UPDATE_MENU_TEXT = "检查更新"

# 手动控制UI管理器
updater = AutoUpdater(main_window)
ui_manager = UpdateUIManager(updater, main_window)
ui_manager.setup_update_menu(menu_bar, "工具(T)")

# 检查状态
print(f"菜单已设置: {ui_manager.is_menu_setup()}")
print(f"菜单动作: {len(ui_manager.get_menu_actions())}")
```

## 架构设计

### 模块结构

```
auto_updater/ui/
├── __init__.py          # 模块导出和便捷函数
├── ui_manager.py        # 主要UI管理器
├── dialogs.py           # 对话框组件
├── widgets.py           # 通用UI组件
├── resources.py         # 文本和样式资源
├── examples.py          # 使用示例
└── README.md            # 本文档
```

### 核心组件

#### UpdateUIManager
主要的UI管理器，负责：
- 菜单设置和管理
- 对话框创建和显示
- 用户交互处理
- 资源清理

#### 对话框组件
- **UpdateProgressDialog**: 更新进度对话框
- **AboutDialog**: 关于对话框
- **UpdateThread**: 后台更新线程

#### 资源管理
- **UpdateUIText**: 文本常量管理
- **UpdateUIStyle**: 样式常量管理
- **UpdateUIConfig**: 配置管理

## API 文档

### UpdateUIManager

#### 构造函数
```python
UpdateUIManager(auto_updater, parent=None)
```
- `auto_updater`: AutoUpdater实例（必需）
- `parent`: 父窗口组件（可选）

#### 主要方法

##### setup_update_menu()
```python
setup_update_menu(menu_bar, menu_title=None) -> None
```
设置更新菜单。
- `menu_bar`: 主窗口菜单栏
- `menu_title`: 菜单标题（可选）

##### check_for_updates_with_ui()
```python
check_for_updates_with_ui(force_check=True) -> None
```
检查更新（带UI交互）。
- `force_check`: 是否强制检查更新

##### show_about_dialog()
```python
show_about_dialog() -> None
```
显示关于对话框。

##### is_menu_setup()
```python
is_menu_setup() -> bool
```
检查菜单是否已设置。

##### get_menu_actions()
```python
get_menu_actions() -> list
```
获取更新相关的菜单动作列表。

##### enable_update_menu()
```python
enable_update_menu(enabled=True) -> None
```
启用或禁用更新菜单。
- `enabled`: 是否启用

##### cleanup()
```python
cleanup() -> None
```
清理所有UI资源。

### 便捷函数

#### create_update_ui()
```python
create_update_ui(updater_instance, parent_window=None) -> UpdateUIManager
```
创建更新UI管理器。

#### setup_standard_update_ui()
```python
setup_standard_update_ui(updater_instance, menu_bar, menu_title=None) -> UpdateUIManager
```
设置标准更新UI。

## 自定义

### 自定义文本

```python
from auto_updater.ui import UpdateUIText

# 修改应用名称
UpdateUIText.APP_NAME = "我的应用程序"

# 修改菜单文本
UpdateUIText.UPDATE_MENU_TEXT = "检查更新"
UpdateUIText.HELP_MENU_TITLE = "工具(T)"

# 修改对话框文本
UpdateUIText.NEW_VERSION_FOUND_TITLE = "发现新版本！"
UpdateUIText.UPDATE_COMPLETE_MESSAGE = "更新完成！"
```

### 自定义样式

```python
from auto_updater.ui import UpdateUIStyle

# 修改对话框尺寸
UpdateUIStyle.PROGRESS_DIALOG_SIZE = QSize(500, 200)
UpdateUIStyle.ABOUT_DIALOG_SIZE = QSize(500, 450)

# 修改颜色方案
UpdateUIStyle.STATUS_COLORS['success'] = '#00ff00'
UpdateUIStyle.STATUS_COLORS['error'] = '#ff0000'
```

### 自定义配置

```python
from auto_updater.ui import UpdateUIConfig

# 修改自动更新设置
UpdateUIConfig.AUTO_CHECK_ON_STARTUP = False
UpdateUIConfig.AUTO_CHECK_INTERVAL = 3600  # 1小时

# 修改UI行为
UpdateUIConfig.SHOW_DETAILED_PROGRESS = True
UpdateUIConfig.ENABLE_ANIMATIONS = False
```

## 移植指南

### 到其他项目

1. **复制模块**：将整个 `auto_updater` 文件夹复制到目标项目
2. **安装依赖**：确保安装了必要的依赖（PyQt5, requests等）
3. **集成代码**：在主窗口中添加几行代码

```python
# 在你的主窗口类中
class MyMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.setup_update_feature()

    def setup_update_feature(self):
        try:
            from auto_updater import AutoUpdater
            from auto_updater.ui import UpdateUIText

            # 自定义应用信息
            UpdateUIText.APP_NAME = "我的应用程序"

            # 初始化更新功能
            self.auto_updater = AutoUpdater(self)
            success = self.auto_updater.setup_update_ui(self.menuBar())

            if success:
                print("更新功能集成成功")
            else:
                print("更新功能集成失败")

        except Exception as e:
            print(f"更新功能设置失败: {e}")
            self.auto_updater = None

    def closeEvent(self, event):
        # 清理资源
        if hasattr(self, 'auto_updater') and self.auto_updater:
            self.auto_updater.cleanup()
        event.accept()
```

### 最小集成要求

- Python 3.7+
- PyQt5
- requests（如果需要网络功能）

## 错误处理

模块提供完善的错误处理机制：

1. **导入错误**：UI组件不可用时自动回退
2. **参数验证**：检查必需参数的有效性
3. **重复设置**：防止重复设置菜单
4. **资源清理**：自动清理UI资源
5. **异常捕获**：捕获并记录所有异常

## 最佳实践

1. **使用标准接口**：优先使用 `updater.setup_update_ui()` 方法
2. **自定义资源**：通过修改 `UpdateUIText` 类来自定义文本
3. **清理资源**：在应用关闭时调用 `cleanup()` 方法
4. **错误处理**：始终检查返回值并处理异常
5. **日志记录**：启用日志记录以便调试

## 故障排除

### 常见问题

1. **导入失败**
   ```
   ImportError: cannot import name 'UpdateUIManager'
   ```
   - 检查模块路径是否正确
   - 确保所有依赖已安装

2. **UI设置失败**
   ```
   更新UI设置失败
   ```
   - 检查父窗口是否有效
   - 确保菜单栏参数不为None

3. **菜单重复**
   - 模块有重复设置防护，不会创建重复菜单
   - 检查是否多次调用了设置方法

### 调试技巧

1. **启用日志**
   ```python
   import logging
   logging.basicConfig(level=logging.DEBUG)
   ```

2. **检查状态**
   ```python
   ui_manager = updater.ui_manager
   print(f"UI可用: {ui_manager is not None}")
   print(f"菜单已设置: {ui_manager.is_menu_setup()}")
   ```

3. **查看资源**
   ```python
   from auto_updater.ui import UpdateUIText
   print(f"应用名称: {UpdateUIText.APP_NAME}")
   ```

## 版本历史

### v1.0.0
- 初始版本
- 完整的UI解耦实现
- 基本功能和错误处理
- 使用示例和文档

## 贡献

欢迎提交问题报告和功能请求。

## 许可证

本项目采用 MIT 许可证。