#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AutoUpdater标准集成示例

这是推荐的集成方式，包含完善的错误处理和日志记录。
"""

import sys
import logging
from PyQt5.QtWidgets import QMainWindow, QApplication, QLabel, QVBoxLayout, QWidget, QTextEdit
from auto_updater import AutoUpdater, UI_AVAILABLE

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)


class StandardExampleWindow(QMainWindow):
    """标准集成示例窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("AutoUpdater 标准集成示例")
        self.setGeometry(100, 100, 700, 500)

        # 初始化变量
        self.auto_updater = None

        # 设置界面
        self.setup_ui()

        # 设置自动更新功能
        self.setup_auto_update()

    def setup_ui(self):
        """设置用户界面"""
        central_widget = QWidget()
        layout = QVBoxLayout()

        # 标题
        title_label = QLabel("<h2>AutoUpdater 标准集成示例</h2>")
        layout.addWidget(title_label)

        # 功能说明
        info_text = QTextEdit()
        info_text.setHtml("""
        <h3>标准集成特点：</h3>
        <ul>
        <li><b>完善的错误处理</b>：捕获和处理各种异常情况</li>
        <li><b>详细日志记录</b>：记录所有重要操作和错误信息</li>
        <li><b>模块化设计</b>：代码结构清晰，易于维护</li>
        <li><b>可配置性</b>：支持各种自定义设置</li>
        </ul>

        <h3>集成代码：</h3>
        <pre style='background-color: #f4f4f4; padding: 10px; border-radius: 5px;'>
def setup_auto_update(self):
    try:
        if UI_AVAILABLE:
            self.auto_updater = AutoUpdater(self)
            success = self.auto_updater.setup_update_ui(
                self.menuBar(), "帮助(H)"
            )

            if success:
                logger.info("自动更新功能集成成功")
                # 可选：自定义配置
                self.auto_updater.enable_auto_check = True
                self.auto_updater.check_interval = 24 * 3600
            else:
                logger.warning("自动更新功能集成失败")
                self.auto_updater = None
        else:
            logger.warning("UI模块不可用")
            self.auto_updater = None

    except Exception as e:
        logger.error(f"自动更新器初始化失败: {e}")
        self.auto_updater = None
        </pre>

        <h3>日志输出：</h3>
        <p>所有的操作日志都会显示在控制台中，便于调试和监控。</p>
        """)
        info_text.setReadOnly(True)
        layout.addWidget(info_text)

        # 状态标签
        self.status_label = QLabel("状态：正在初始化...")
        self.status_label.setStyleSheet("color: blue; font-weight: bold;")
        layout.addWidget(self.status_label)

        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # 创建菜单栏
        self.create_menus()

    def create_menus(self):
        """创建菜单栏"""
        # 检查更新菜单项（手动添加，用于演示）
        help_menu = self.menuBar().addMenu("演示(&D)")
        check_action = help_menu.addAction("检查集成状态")
        check_action.triggered.connect(self.check_integration_status)

    def setup_auto_update(self):
        """标准集成方案"""
        try:
            logger.info("开始初始化自动更新功能...")

            if UI_AVAILABLE:
                # 初始化更新器
                self.auto_updater = AutoUpdater(self)
                logger.info("AutoUpdater实例创建成功")

                # 设置更新UI
                success = self.auto_updater.setup_update_ui(
                    self.menuBar(), "帮助(H)"
                )

                if success:
                    logger.info("自动更新功能集成成功")
                    self.status_label.setText("状态：自动更新功能已成功集成")
                    self.status_label.setStyleSheet("color: green; font-weight: bold;")

                    # 自定义配置
                    self.auto_updater.enable_auto_check = True
                    self.auto_updater.check_interval = 24 * 3600  # 24小时
                    logger.info("自动更新配置完成：启用自动检查，间隔24小时")

                else:
                    logger.warning("自动更新功能集成失败")
                    self.status_label.setText("状态：自动更新功能集成失败")
                    self.status_label.setStyleSheet("color: orange; font-weight: bold;")
                    self.auto_updater = None

            else:
                logger.warning("UI模块不可用，跳过自动更新功能")
                self.status_label.setText("状态：UI模块不可用")
                self.status_label.setStyleSheet("color: red; font-weight: bold;")
                self.auto_updater = None

        except ImportError as e:
            logger.error(f"导入自动更新模块失败: {e}")
            self.status_label.setText(f"状态：模块导入失败 - {e}")
            self.status_label.setStyleSheet("color: red; font-weight: bold;")
            self.auto_updater = None

        except Exception as e:
            logger.error(f"自动更新器初始化失败: {e}")
            self.status_label.setText(f"状态：初始化失败 - {e}")
            self.status_label.setStyleSheet("color: red; font-weight: bold;")
            self.auto_updater = None

    def check_integration_status(self):
        """检查集成状态"""
        from PyQt5.QtWidgets import QMessageBox

        status_info = f"""
        自动更新功能集成状态检查：

        <b>AutoUpdater实例：</b> {'✅ 已创建' if self.auto_updater else '❌ 未创建'}
        <b>UI模块可用性：</b> {'✅ 可用' if UI_AVAILABLE else '❌ 不可用'}

        <b>配置信息：</b>
        - 应用名称: {getattr(self.auto_updater, 'app_name', '未知') if self.auto_updater else 'N/A'}
        - 当前版本: {getattr(self.auto_updater, 'current_version', '未知') if self.auto_updater else 'N/A'}
        - GitHub仓库: {getattr(self.auto_updater, 'github_repo', '未知') if self.auto_updater else 'N/A'}

        <b>功能状态：</b>
        - 自动检查: {'启用' if hasattr(self.auto_updater, 'enable_auto_check') and self.auto_updater.enable_auto_check else '禁用'}
        - 检查间隔: {getattr(self.auto_updater, 'check_interval', '未知') if self.auto_updater else 'N/A'}秒
        """

        QMessageBox.information(self, "集成状态", status_info)

    def closeEvent(self, event):
        """应用退出时清理资源"""
        logger.info("应用程序正在退出...")

        if hasattr(self, 'auto_updater') and self.auto_updater:
            try:
                self.auto_updater.cleanup()
                logger.info("自动更新器资源已清理")
            except Exception as e:
                logger.error(f"清理自动更新器资源时出错: {e}")

        logger.info("应用程序退出")
        event.accept()


def main():
    """主函数"""
    # 创建应用
    app = QApplication(sys.argv)
    app.setApplicationName("AutoUpdater标准集成示例")

    # 创建主窗口
    window = StandardExampleWindow()
    window.show()

    # 输出启动信息
    logger.info("应用程序启动完成")

    # 运行应用
    exit_code = app.exec_()
    logger.info(f"应用程序退出，退出代码: {exit_code}")

    return exit_code


if __name__ == "__main__":
    sys.exit(main())