# -*- coding: utf-8 -*-
"""
auto_updater.ui 子包
提供独立的自动更新功能UI组件

这个包实现了自动更新功能的完整UI层，与业务逻辑完全解耦。

主要组件：
- UpdateUIManager: 主要的UI管理器，提供统一的接口
- UpdateProgressDialog: 更新进度对话框
- AboutDialog: 关于对话框
- UpdateThread: 更新操作的后台线程
- UpdateStatusWidget: 更新状态显示组件
- UpdateUIText: UI文本常量
- UpdateUIStyle: UI样式常量

使用示例：
    # 基本使用
    from auto_updater import AutoUpdater
    from auto_updater.ui import UpdateUIManager

    # 在主窗口中
    updater = AutoUpdater(main_window)
    ui_manager = UpdateUIManager(updater, main_window)
    ui_manager.setup_update_menu(menu_bar)

    # 或者使用便捷方法
    updater.setup_update_ui(menu_bar)

特点：
- 完全解耦的UI组件
- 统一的资源管理
- 完善的错误处理
- 支持自定义和扩展
"""

from .ui_manager import UpdateUIManager
from .dialogs import UpdateProgressDialog, AboutDialog, UpdateThread
from .widgets import UpdateStatusWidget
from .resources import UpdateUIText, UpdateUIStyle, UpdateUIConfig

# 版本信息
__version__ = "1.0.0"
__author__ = "Auto Update UI Module"

# 导出的公共接口
__all__ = [
    # 主要组件
    'UpdateUIManager',
    'UpdateProgressDialog',
    'AboutDialog',
    'UpdateThread',
    'UpdateStatusWidget',

    # 资源和配置
    'UpdateUIText',
    'UpdateUIStyle',
    'UpdateUIConfig',

    # 版本信息
    '__version__',
    '__author__'
]

# 模块级别的便利函数
def create_update_ui(updater_instance, parent_window=None):
    """
    便捷函数：创建更新UI管理器

    Args:
        updater_instance: AutoUpdater实例
        parent_window: 父窗口，可选

    Returns:
        UpdateUIManager: UI管理器实例

    Example:
        updater = AutoUpdater()
        ui_manager = create_update_ui(updater, main_window)
        ui_manager.setup_update_menu(menu_bar)
    """
    return UpdateUIManager(updater_instance, parent_window)

def setup_standard_update_ui(updater_instance, menu_bar, menu_title=None):
    """
    便捷函数：设置标准更新UI

    Args:
        updater_instance: AutoUpdater实例
        menu_bar: 主窗口菜单栏
        menu_title: 菜单标题，可选

    Returns:
        UpdateUIManager: UI管理器实例

    Example:
        updater = AutoUpdater()
        ui_manager = setup_standard_update_ui(updater, menu_bar)
    """
    ui_manager = UpdateUIManager(updater_instance, menu_bar.parentWidget() if menu_bar else None)
    ui_manager.setup_update_menu(menu_bar, menu_title)
    return ui_manager