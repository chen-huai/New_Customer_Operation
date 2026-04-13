# -*- coding: utf-8 -*-
"""
自动更新UI模块使用示例

这个文件展示了如何在不同的项目中使用auto_updater.ui模块
"""

# 示例1：基本使用
def basic_usage_example():
    """
    基本使用示例：最简单的集成方式
    """
    from PyQt5.QtWidgets import QMainWindow, QApplication, QMenuBar
    from auto_updater import AutoUpdater

    # 创建主窗口
    app = QApplication([])
    main_window = QMainWindow()
    menu_bar = main_window.menuBar()

    # 初始化更新器
    updater = AutoUpdater(main_window)

    # 一行代码完成所有UI设置
    success = updater.setup_update_ui(menu_bar)

    if success:
        print("更新UI设置成功！")
    else:
        print("更新UI设置失败")

    main_window.show()
    app.exec_()

# 示例2：手动控制UI管理器
def manual_control_example():
    """
    手动控制示例：更精细的控制方式
    """
    from PyQt5.QtWidgets import QMainWindow, QApplication, QMenuBar
    from auto_updater import AutoUpdater
    from auto_updater.ui import UpdateUIManager

    # 创建主窗口
    app = QApplication([])
    main_window = QMainWindow()
    menu_bar = main_window.menuBar()

    # 初始化组件
    updater = AutoUpdater(main_window)
    ui_manager = UpdateUIManager(updater, main_window)

    # 手动设置菜单
    try:
        ui_manager.setup_update_menu(menu_bar, "工具(T)")
        print("菜单设置成功")

        # 检查状态
        print(f"菜单是否已设置: {ui_manager.is_menu_setup()}")

        # 获取菜单动作
        actions = ui_manager.get_menu_actions()
        print(f"创建了 {len(actions)} 个菜单动作")

        # 控制菜单启用状态
        ui_manager.enable_update_menu(True)
        print("更新菜单已启用")

    except Exception as e:
        print(f"设置失败: {e}")

    main_window.show()
    app.exec_()

# 示例3：自定义UI样式
def custom_style_example():
    """
    自定义样式示例：修改UI文本和样式
    """
    from PyQt5.QtWidgets import QMainWindow, QApplication, QMenuBar
    from auto_updater import AutoUpdater
    from auto_updater.ui import UpdateUIManager, UpdateUIText, UpdateUIStyle

    # 自定义文本
    UpdateUIText.APP_NAME = "我的应用程序"
    UpdateUIText.UPDATE_MENU_TEXT = "检查更新"
    UpdateUIText.HELP_MENU_TITLE = "选项(O)"

    # 创建主窗口
    app = QApplication([])
    main_window = QMainWindow()
    menu_bar = main_window.menuBar()

    # 初始化并设置
    updater = AutoUpdater(main_window)
    ui_manager = UpdateUIManager(updater, main_window)
    ui_manager.setup_update_menu(menu_bar)

    print("自定义样式UI设置完成")

    main_window.show()
    app.exec_()

# 示例4：移植到其他项目
def portable_integration_example():
    """
    移植集成示例：如何在其他项目中使用
    """
    """
    假设你有一个现有的PyQt5应用程序，想要集成更新功能：

    1. 将auto_updater文件夹复制到你的项目中
    2. 在你的主窗口类中添加以下代码：

    class MyMainWindow(QMainWindow):
        def __init__(self):
            super().__init__()
            self.setup_ui()
            self.setup_update_feature()

        def setup_update_feature(self):
            # 集成更新功能
            try:
                from auto_updater import AutoUpdater
                from auto_updater.ui import UpdateUIText

                # 自定义应用信息
                UpdateUIText.APP_NAME = "我的应用程序"

                # 初始化更新器
                self.auto_updater = AutoUpdater(self)
                success = self.auto_updater.setup_update_ui(self.menuBar())

                if success:
                    print("更新功能集成成功")
                else:
                    print("更新功能集成失败")

            except ImportError:
                print("无法导入更新模块")
                self.auto_updater = None
            except Exception as e:
                print(f"更新功能设置失败: {e}")
                self.auto_updater = None

        def closeEvent(self, event):
            # 清理资源
            if hasattr(self, 'auto_updater') and self.auto_updater:
                self.auto_updater.cleanup()
            event.accept()

    3. 就这样！你的应用现在有了完整的自动更新功能
    """

# 示例5：错误处理和回退
def error_handling_example():
    """
    错误处理示例：展示如何处理各种异常情况
    """
    from PyQt5.QtWidgets import QMainWindow, QApplication, QMenuBar, QMessageBox
    from auto_updater import AutoUpdater

    class MainWindow(QMainWindow):
        def __init__(self):
            super().__init__()
            self.setWindowTitle("错误处理示例")
            self.menu_bar = self.menuBar()
            self.auto_updater = None
            self.setup_update_with_error_handling()

        def setup_update_with_error_handling(self):
            """带错误处理的更新功能设置"""
            try:
                # 尝试初始化更新器
                self.auto_updater = AutoUpdater(self)

                # 尝试设置UI
                success = self.auto_updater.setup_update_ui(self.menu_bar)

                if success:
                    self.statusBar().showMessage("更新功能已启用", 3000)
                else:
                    self.show_fallback_message()

            except ImportError as e:
                self.show_error_dialog("模块导入失败", f"无法导入更新模块: {e}")
            except RuntimeError as e:
                self.show_error_dialog("运行时错误", f"更新功能设置失败: {e}")
            except Exception as e:
                self.show_error_dialog("未知错误", f"发生未知错误: {e}")

        def show_error_dialog(self, title, message):
            """显示错误对话框"""
            QMessageBox.warning(self, title, message)

        def show_fallback_message(self):
            """显示回退消息"""
            self.statusBar().showMessage("更新功能不可用", 5000)

        def closeEvent(self, event):
            """清理资源"""
            if self.auto_updater:
                try:
                    self.auto_updater.cleanup()
                except Exception as e:
                    print(f"清理资源时出错: {e}")
            event.accept()

    # 运行示例
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()

# 示例6：高级自定义
def advanced_customization_example():
    """
    高级自定义示例：深度定制更新功能
    """
    from PyQt5.QtWidgets import QMainWindow, QApplication, QMenuBar, QAction
    from auto_updater import AutoUpdater
    from auto_updater.ui import UpdateUIManager, UpdateUIText

    class CustomUpdateUIManager(UpdateUIManager):
        """自定义UI管理器"""

        def setup_update_menu(self, menu_bar, menu_title=None):
            """重写菜单设置方法，使用自定义布局"""
            try:
                # 创建专门的更新菜单
                update_menu = menu_bar.addMenu("更新(&U)")

                # 检查更新动作
                check_action = QAction("检查更新...", self.parent)
                check_action.triggered.connect(self.check_for_updates_with_ui)
                check_action.setShortcut("F5")  # 添加快捷键
                update_menu.addAction(check_action)

                # 分隔符
                update_menu.addSeparator()

                # 关于动作
                about_action = QAction("关于", self.parent)
                about_action.triggered.connect(self.show_about_dialog)
                update_menu.addAction(about_action)

                # 添加到父菜单的帮助菜单中
                help_menu = menu_bar.addMenu("帮助(&H)")
                help_menu.addSeparator()
                help_menu.addAction(check_action)
                help_menu.addAction(about_action)

                self._is_menu_setup = True
                print("自定义更新菜单设置完成")

            except Exception as e:
                print(f"自定义菜单设置失败: {e}")
                raise

    # 使用自定义UI管理器
    app = QApplication([])
    main_window = QMainWindow()
    menu_bar = main_window.menuBar()

    updater = AutoUpdater(main_window)
    custom_ui = CustomUpdateUIManager(updater, main_window)
    custom_ui.setup_update_menu(menu_bar)

    main_window.show()
    app.exec_()

# 主函数：运行所有示例
def main():
    """
    主函数：运行所有示例
    你可以取消注释想要运行的示例
    """
    import sys

    print("=== 自动更新UI模块使用示例 ===")
    print("请选择要运行的示例：")
    print("1. 基本使用示例")
    print("2. 手动控制示例")
    print("3. 自定义样式示例")
    print("4. 错误处理示例")
    print("5. 高级自定义示例")

    # 取消注释想要运行的示例
    # basic_usage_example()
    # manual_control_example()
    # custom_style_example()
    # error_handling_example()
    # advanced_customization_example()

    print("请取消注释相应的示例函数来运行")

if __name__ == "__main__":
    main()