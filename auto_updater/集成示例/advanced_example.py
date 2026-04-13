#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AutoUpdater深度定制集成示例

这是完全自定义的集成方式，展示如何完全控制更新流程和用户界面。
"""

import sys
import logging
import time
from PyQt5.QtWidgets import (
    QMainWindow, QApplication, QLabel, QVBoxLayout, QWidget,
    QPushButton, QTextEdit, QMessageBox, QProgressBar, QHBoxLayout
)
from PyQt5.QtCore import QThread, pyqtSignal, QTimer
from auto_updater import AutoUpdater, UpdateUIManager
from auto_updater.ui.dialogs import UpdateProgressDialog

# 配置详细日志
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)


class UpdateCheckThread(QThread):
    """更新检查线程"""

    update_found = pyqtSignal(str, str)  # remote_version, local_version
    no_update = pyqtSignal(str)          # local_version
    error_occurred = pyqtSignal(str)     # error_message

    def __init__(self, updater):
        super().__init__()
        self.updater = updater

    def run(self):
        """执行更新检查"""
        try:
            logger.info("开始检查更新...")

            has_update, remote_version, local_version, error = \
                self.updater.check_for_updates(force_check=True)

            if error:
                logger.error(f"更新检查失败: {error}")
                self.error_occurred.emit(error)
            elif has_update:
                logger.info(f"发现新版本: {remote_version}")
                self.update_found.emit(remote_version, local_version)
            else:
                logger.info(f"已是最新版本: {local_version}")
                self.no_update.emit(local_version)

        except Exception as e:
            logger.error(f"更新检查异常: {e}")
            self.error_occurred.emit(str(e))


class AdvancedExampleWindow(QMainWindow):
    """深度定制集成示例窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("AutoUpdater 深度定制集成示例")
        self.setGeometry(100, 100, 800, 600)

        # 初始化变量
        self.auto_updater = None
        self.update_ui_manager = None
        self.update_thread = None

        # 设置界面
        self.setup_ui()

        # 设置自定义自动更新
        self.setup_custom_auto_update()

        # 设置状态更新定时器
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self.update_status_display)
        self.status_timer.start(2000)  # 每2秒更新一次状态

    def setup_ui(self):
        """设置用户界面"""
        central_widget = QWidget()
        layout = QVBoxLayout()

        # 标题
        title_label = QLabel("<h2>AutoUpdater 深度定制集成示例</h2>")
        layout.addWidget(title_label)

        # 功能说明区域
        info_text = QTextEdit()
        info_text.setHtml("""
        <h3>深度定制特点：</h3>
        <ul>
        <li><b>完全自定义UI</b>：自定义所有用户界面和交互逻辑</li>
        <li><b>异步操作</b>：在后台线程中执行更新检查，不阻塞主界面</li>
        <li><b>自定义对话框</b>：完全控制更新确认和进度显示</li>
        <li><b>事件驱动</b>：使用信号槽机制处理各种事件</li>
        <li><b>详细状态显示</b>：实时显示更新状态和详细信息</li>
        </ul>

        <h3>自定义功能演示：</h3>
        <ul>
        <li>点击下方按钮测试各种自定义功能</li>
        <li>查看实时状态显示了解更新流程</li>
        <li>注意异步操作不会阻塞界面</li>
        </ul>
        """)
        info_text.setReadOnly(True)
        info_text.setMaximumHeight(200)
        layout.addWidget(info_text)

        # 控制按钮区域
        button_layout = QHBoxLayout()

        # 检查更新按钮
        self.check_button = QPushButton("🔍 检查更新")
        self.check_button.clicked.connect(self.custom_check_updates)
        button_layout.addWidget(self.check_button)

        # 获取版本信息按钮
        self.version_button = QPushButton("📋 版本信息")
        self.version_button.clicked.connect(self.get_version_info)
        button_layout.addWidget(self.version_button)

        # 测试网络连接按钮
        self.network_button = QPushButton("🌐 测试网络")
        self.network_button.clicked.connect(self.test_network_connection)
        button_layout.addWidget(self.network_button)

        layout.addLayout(button_layout)

        # 进度条
        progress_layout = QHBoxLayout()
        progress_layout.addWidget(QLabel("操作进度:"))
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        progress_layout.addWidget(self.progress_bar)
        layout.addLayout(progress_layout)

        # 状态显示区域
        self.status_text = QTextEdit()
        self.status_text.setMaximumHeight(150)
        self.status_text.setHtml("<b>状态：</b>正在初始化...")
        self.status_text.setReadOnly(True)
        layout.addWidget(self.status_text)

        # 详细日志区域
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(100)
        self.log_text.setPlaceholderText("详细日志将显示在这里...")
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # 创建自定义菜单
        self.create_custom_menus()

    def create_custom_menus(self):
        """创建自定义菜单"""
        # 演示菜单
        demo_menu = self.menuBar().addMenu("演示(&D)")

        # 检查更新
        check_action = demo_menu.addAction("自定义检查更新")
        check_action.triggered.connect(self.custom_check_updates)

        demo_menu.addSeparator()

        # 获取版本信息
        version_action = demo_menu.addAction("获取版本信息")
        version_action.triggered.connect(self.get_version_info)

        demo_menu.addSeparator()

        # 测试网络
        network_action = demo_menu.addAction("测试网络连接")
        network_action.triggered.connect(self.test_network_connection)

        # 帮助菜单
        help_menu = self.menuBar().addMenu("帮助(&H)")

        # 关于
        about_action = help_menu.addAction("关于")
        about_action.triggered.connect(self.show_custom_about)

    def setup_custom_auto_update(self):
        """深度定制集成方案"""
        try:
            self.log_message("开始初始化自定义自动更新功能...")

            # 初始化更新器
            self.auto_updater = AutoUpdater(self)
            self.log_message("AutoUpdater实例创建成功")

            # 创建独立的UI管理器
            self.update_ui_manager = UpdateUIManager(
                self.auto_updater, self
            )
            self.log_message("UI管理器创建成功")

            # 自定义配置
            self.configure_update_settings()

            self.log_message("自定义自动更新功能初始化完成")
            self.update_status("状态：自动更新功能已成功集成", "green")

        except Exception as e:
            self.log_message(f"自定义自动更新器初始化失败: {e}", "ERROR")
            self.update_status("状态：初始化失败", "red")
            self.auto_updater = None
            self.update_ui_manager = None

    def configure_update_settings(self):
        """配置更新设置"""
        if self.auto_updater:
            # 自定义设置
            self.auto_updater.enable_auto_check = True
            self.auto_updater.check_interval = 24 * 3600  # 24小时

            self.log_message("更新设置配置完成")
            self.log_message(f"- 自动检查: {self.auto_updater.enable_auto_check}")
            self.log_message(f"- 检查间隔: {self.auto_updater.check_interval}秒")

    def custom_check_updates(self):
        """自定义更新检查流程"""
        if not self.auto_updater:
            self.show_error("自动更新功能不可用")
            return

        # 显示检查状态
        self.update_status("状态：正在检查更新...", "blue")
        self.show_progress(True)
        self.log_message("开始执行自定义更新检查...")

        # 使用异步线程检查更新
        self.update_thread = UpdateCheckThread(self.auto_updater)
        self.update_thread.update_found.connect(self.on_update_found)
        self.update_thread.no_update.connect(self.on_no_update)
        self.update_thread.error_occurred.connect(self.on_update_error)
        self.update_thread.finished.connect(lambda: self.show_progress(False))
        self.update_thread.start()

    def on_update_found(self, remote_version, local_version):
        """发现更新时的回调"""
        self.log_message(f"发现新版本: {remote_version} (当前: {local_version})")
        self.update_status(f"状态：发现新版本 {remote_version}", "orange")

        # 自定义更新确认对话框
        reply = QMessageBox.question(
            self, "发现新版本",
            f"🚀 发现新版本！\n\n"
            f"当前版本: {local_version}\n"
            f"最新版本: {remote_version}\n\n"
            "是否立即下载更新？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )

        if reply == QMessageBox.Yes:
            self.show_custom_update_dialog(remote_version, local_version)

    def on_no_update(self, local_version):
        """没有更新时的回调"""
        self.log_message(f"已是最新版本: {local_version}")
        self.update_status(f"状态：已是最新版本 {local_version}", "green")
        self.show_info(f"恭喜！您的应用程序已是最新版本 {local_version}")

    def on_update_error(self, error_message):
        """更新检查错误时的回调"""
        self.log_message(f"更新检查失败: {error_message}", "ERROR")
        self.update_status("状态：更新检查失败", "red")
        self.show_error(f"更新检查失败: {error_message}")

    def show_custom_update_dialog(self, remote_version, local_version):
        """显示自定义更新对话框"""
        self.log_message("显示自定义更新对话框")

        # 创建自定义进度对话框
        dialog = UpdateProgressDialog(self, remote_version, local_version)

        # 连接下载完成回调
        dialog.download_finished.connect(
            lambda success, msg: self.on_download_finished(success, msg, dialog)
        )

        # 开始更新流程
        self.log_message(f"开始下载更新 {remote_version}")
        dialog.start_update(remote_version, self.auto_updater)

    def on_download_finished(self, success, message, dialog):
        """下载完成回调"""
        dialog.close()
        self.log_message(f"下载完成 - 成功: {success}, 消息: {message}")

        if success:
            self.update_status("状态：更新下载完成", "blue")
            reply = QMessageBox.question(
                self, "更新下载完成",
                "✅ 更新已下载完成！\n\n"
                "是否立即安装？\n"
                "⚠️ 安装后将自动重启应用程序。",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )

            if reply == QMessageBox.Yes:
                self.log_message("用户确认安装更新")
                self.show_info("准备安装更新...\n应用程序将在安装完成后自动重启。")
                # 这里可以添加自定义的安装逻辑
        else:
            self.update_status("状态：更新下载失败", "red")
            self.show_error(f"更新下载失败: {message}")

    def get_version_info(self):
        """获取版本信息"""
        if not self.auto_updater:
            self.show_error("自动更新功能不可用")
            return

        version_info = f"""
        <h3>应用程序版本信息</h3>

        <table style='border-collapse: collapse; width: 100%;'>
        <tr style='background-color: #f0f0f0;'>
        <td style='padding: 8px; border: 1px solid #ddd;'><b>配置项</b></td>
        <td style='padding: 8px; border: 1px solid #ddd;'><b>值</b></td>
        </tr>
        <tr>
        <td style='padding: 8px; border: 1px solid #ddd;'>应用名称</td>
        <td style='padding: 8px; border: 1px solid #ddd;'>{getattr(self.auto_updater, 'app_name', '未知')}</td>
        </tr>
        <tr>
        <td style='padding: 8px; border: 1px solid #ddd;'>当前版本</td>
        <td style='padding: 8px; border: 1px solid #ddd;'>{getattr(self.auto_updater, 'current_version', '未知')}</td>
        </tr>
        <tr>
        <td style='padding: 8px; border: 1px solid #ddd;'>GitHub仓库</td>
        <td style='padding: 8px; border: 1px solid #ddd;'>{getattr(self.auto_updater, 'github_repo', '未知')}</td>
        </tr>
        <tr>
        <td style='padding: 8px; border: 1px solid #ddd;'>自动检查</td>
        <td style='padding: 8px; border: 1px solid #ddd;'>{'启用' if getattr(self.auto_updater, 'enable_auto_check', False) else '禁用'}</td>
        </tr>
        <tr>
        <td style='padding: 8px; border: 1px solid #ddd;'>检查间隔</td>
        <td style='padding: 8px; border: 1px solid #ddd;'>{getattr(self.auto_updater, 'check_interval', '未知')}秒</td>
        </tr>
        </table>
        """

        QMessageBox.information(self, "版本信息", version_info)
        self.log_message("获取版本信息完成")

    def test_network_connection(self):
        """测试网络连接"""
        self.log_message("开始测试网络连接...")
        self.update_status("状态：正在测试网络连接...", "blue")
        self.show_progress(True)

        try:
            # 测试GitHub API连接
            import requests
            response = requests.get('https://api.github.com', timeout=10)

            if response.status_code == 200:
                self.log_message("GitHub API连接正常")
                self.update_status("状态：网络连接正常", "green")
                self.show_info("✅ 网络连接测试成功！\nGitHub API可正常访问。")
            else:
                self.log_message(f"GitHub API连接异常: HTTP {response.status_code}")
                self.update_status(f"状态：网络连接异常 (HTTP {response.status_code})", "orange")
                self.show_warning("⚠️ 网络连接异常\nGitHub API返回状态码: " + str(response.status_code))

        except requests.exceptions.Timeout:
            self.log_message("网络连接超时", "ERROR")
            self.update_status("状态：网络连接超时", "red")
            self.show_error("❌ 网络连接超时\n请检查网络设置或稍后重试。")

        except requests.exceptions.ConnectionError:
            self.log_message("网络连接失败", "ERROR")
            self.update_status("状态：网络连接失败", "red")
            self.show_error("❌ 网络连接失败\n请检查网络连接状态。")

        except Exception as e:
            self.log_message(f"网络测试异常: {e}", "ERROR")
            self.update_status("状态：网络测试异常", "red")
            self.show_error(f"❌ 网络测试异常: {e}")

        finally:
            self.show_progress(False)

    def show_custom_about(self):
        """自定义关于对话框"""
        about_text = f"""
        <h2>AutoUpdater 深度定制集成示例</h2>

        <p><b>版本:</b> {self.auto_updater.current_version if self.auto_updater else '未知'}</p>
        <p><b>自动更新功能:</b> ✅ 已集成（深度定制）</p>
        <p><b>集成方式:</b> 完全自定义UI和逻辑</p>

        <h3>功能特性:</h3>
        <ul>
        <li>✅ 异步更新检查</li>
        <li>✅ 自定义用户界面</li>
        <li>✅ 实时状态显示</li>
        <li>✅ 详细的日志记录</li>
        <li>✅ 网络连接测试</li>
        <li>✅ 完整的错误处理</li>
        </ul>

        <p><b>© 2025 AutoUpdater开发团队</b></p>
        <p>这个示例展示了如何完全自定义自动更新功能的集成方式。</p>
        """

        QMessageBox.about(self, "关于", about_text)
        self.log_message("显示自定义关于对话框")

    def update_status_display(self):
        """定时更新状态显示"""
        if self.auto_updater:
            # 这里可以添加实时状态更新逻辑
            pass

    def show_progress(self, show):
        """显示或隐藏进度条"""
        self.progress_bar.setVisible(show)
        if show:
            self.progress_bar.setRange(0, 0)  # 不确定进度
        else:
            self.progress_bar.setRange(0, 1)  # 停止动画

    def update_status(self, message, color="black"):
        """更新状态显示"""
        self.status_text.setHtml(f"<b>状态：</b><span style='color: {color}; font-weight: bold;'>{message}</span>")
        self.log_message(f"状态更新: {message}")

    def log_message(self, message, level="INFO"):
        """添加日志消息"""
        timestamp = time.strftime("%H:%M:%S")

        if level == "ERROR":
            color = "red"
            icon = "❌"
        elif level == "WARNING":
            color = "orange"
            icon = "⚠️"
        else:
            color = "blue"
            icon = "ℹ️"

        log_entry = f"<span style='color: gray;'>[{timestamp}]</span> {icon} {message}"
        self.log_text.append(log_entry)

        # 自动滚动到底部
        cursor = self.log_text.textCursor()
        cursor.movePosition(cursor.End)
        self.log_text.setTextCursor(cursor)

    def show_info(self, message):
        """显示信息对话框"""
        QMessageBox.information(self, "信息", message)

    def show_warning(self, message):
        """显示警告对话框"""
        QMessageBox.warning(self, "警告", message)

    def show_error(self, message):
        """显示错误对话框"""
        QMessageBox.critical(self, "错误", message)

    def closeEvent(self, event):
        """自定义清理流程"""
        self.log_message("应用程序正在退出...")

        # 停止状态定时器
        if hasattr(self, 'status_timer'):
            self.status_timer.stop()

        # 等待更新线程完成
        if self.update_thread and self.update_thread.isRunning():
            self.log_message("等待更新线程完成...")
            self.update_thread.quit()
            self.update_thread.wait(3000)  # 最多等待3秒

        # 清理自动更新器资源
        if hasattr(self, 'auto_updater') and self.auto_updater:
            try:
                if hasattr(self, 'update_ui_manager'):
                    self.update_ui_manager.cleanup()
                    self.log_message("UI管理器资源已清理")

                self.auto_updater.cleanup()
                self.log_message("自动更新器资源已清理")

            except Exception as e:
                self.log_message(f"清理自动更新器资源时出错: {e}", "ERROR")

        self.log_message("应用程序退出")
        event.accept()


def main():
    """主函数"""
    # 创建应用
    app = QApplication(sys.argv)
    app.setApplicationName("AutoUpdater深度定制集成示例")

    # 创建主窗口
    window = AdvancedExampleWindow()
    window.show()

    # 输出启动信息
    logger.info("深度定制集成示例启动完成")

    # 运行应用
    exit_code = app.exec_()
    logger.info(f"应用程序退出，退出代码: {exit_code}")

    return exit_code


if __name__ == "__main__":
    sys.exit(main())