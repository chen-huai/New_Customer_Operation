# -*- coding: utf-8 -*-
"""
更新功能对话框组件
包含更新进度对话框、关于对话框和更新线程
"""
import logging
import os
import sys
import subprocess
import webbrowser
from typing import Optional, Callable
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                             QProgressBar, QMessageBox, QFrame, QApplication)
from PyQt5.QtCore import QThread, pyqtSignal, QTimer, Qt
from PyQt5.QtGui import QFont

from .resources import UpdateUIText, UpdateUIStyle
from ..config import get_app_executable_path, is_production_environment
from ..config_constants import APP_EXECUTABLE, APP_NAME

logger = logging.getLogger(__name__)

class UpdateThread(QThread):
    """
    更新线程

    在后台执行下载和安装更新操作，避免阻塞UI线程
    """

    # 信号定义
    progress_signal = pyqtSignal(int, int, int)  # 已下载, 总量, 百分比
    status_signal = pyqtSignal(str)  # 状态信息
    finished_signal = pyqtSignal(bool, str, str)  # 是否成功, 错误信息, 下载路径

    def __init__(self, auto_updater, version: str):
        """
        初始化更新线程

        Args:
            auto_updater: AutoUpdater实例
            version: 要更新的版本号
        """
        super().__init__()
        self.auto_updater = auto_updater
        self.version = version
        self._is_cancelled = False

    def run(self) -> None:
        """执行更新过程"""
        try:
            if self._is_cancelled:
                return

            self.status_signal.emit(UpdateUIText.DOWNLOADING_UPDATE_MESSAGE)

            # 下载更新文件
            success, download_path, error = self.auto_updater.download_update(
                self.version,
                self.progress_callback
            )

            if self._is_cancelled:
                return

            if not success:
                self.finished_signal.emit(False, f"{UpdateUIText.DOWNLOAD_FAILED_MESSAGE}: {error}", None)
                return

            self.status_signal.emit(UpdateUIText.INSTALLING_UPDATE_MESSAGE)

            # 执行更新
            success, error = self.auto_updater.execute_update(download_path, self.version)

            if self._is_cancelled:
                return

            if success:
                self.finished_signal.emit(True, None, download_path)
            else:
                self.finished_signal.emit(False, f"{UpdateUIText.INSTALL_FAILED_MESSAGE}: {error}", download_path)

        except Exception as e:
            if not self._is_cancelled:
                self.finished_signal.emit(False, f"{UpdateUIText.UPDATE_PROCESS_ERROR_MESSAGE}: {str(e)}", None)

    def progress_callback(self, downloaded: int, total: int, percentage: int) -> None:
        """进度回调函数"""
        if not self._is_cancelled:
            self.progress_signal.emit(downloaded, total, percentage)

    def cancel(self) -> None:
        """取消更新"""
        self._is_cancelled = True


class UpdateProgressDialog(QDialog):
    """
    更新进度对话框

    显示更新下载和安装进度，提供用户反馈
    """

    def __init__(self, parent=None):
        """
        初始化更新进度对话框

        Args:
            parent: 父窗口
        """
        super().__init__(parent)
        self.setWindowTitle(UpdateUIText.UPDATING_TITLE)
        self.setFixedSize(UpdateUIStyle.PROGRESS_DIALOG_SIZE)
        self.setModal(True)
        self.update_thread = None
        self.auto_updater = None  # 保存auto_updater引用

        # 添加更新文件路径跟踪
        self.updated_executable_path = None
        self.download_path = None

        self._setup_ui()
        logger.debug("UpdateProgressDialog 初始化完成")

    def _setup_ui(self) -> None:
        """设置UI界面"""
        layout = QVBoxLayout()

        # 状态标签
        self.status_label = QLabel(UpdateUIText.PREPARING_UPDATE_MESSAGE)
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setWordWrap(True)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)

        # 进度标签
        self.progress_label = QLabel("")
        self.progress_label.setAlignment(Qt.AlignCenter)
        self.progress_label.setStyleSheet(UpdateUIStyle.PROGRESS_LABEL_STYLE)

        # 取消按钮
        self.cancel_btn = QPushButton(UpdateUIText.CANCEL_BUTTON_TEXT)
        self.cancel_btn.clicked.connect(self.cancel_update)
        self.cancel_btn.setEnabled(False)  # 开始后不允许取消

        # 添加控件到布局
        layout.addWidget(self.status_label)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.progress_label)
        layout.addWidget(self.cancel_btn)

        self.setLayout(layout)

    def start_update(self, version: str, auto_updater) -> None:
        """
        开始更新

        Args:
            version: 要更新的版本号
            auto_updater: AutoUpdater实例
        """
        try:
            # ✅ 保存auto_updater引用
            self.auto_updater = auto_updater

            self.status_label.setText(f"{UpdateUIText.UPDATING_TO_VERSION_MESSAGE} {version}...")
            self.progress_bar.setValue(0)
            self.cancel_btn.setEnabled(False)
            self.progress_label.setText("")

            # 创建并启动更新线程
            self.update_thread = UpdateThread(auto_updater, version)
            self.update_thread.progress_signal.connect(self._update_progress)
            self.update_thread.status_signal.connect(self._update_status)
            self.update_thread.finished_signal.connect(self._update_finished)
            self.update_thread.start()

        except Exception as e:
            logger.error(f"启动更新失败: {e}")
            QMessageBox.critical(self, UpdateUIText.ERROR_TITLE, f"{UpdateUIText.START_UPDATE_ERROR_MESSAGE}: {str(e)}")
            self.reject()

    def _update_progress(self, downloaded: int, total: int, percentage: int) -> None:
        """
        更新进度

        Args:
            downloaded: 已下载字节数
            total: 总字节数
            percentage: 百分比
        """
        # 数据类型验证和边界检查
        try:
            # 确保数据为数值类型
            downloaded = int(downloaded) if downloaded is not None else 0
            total = int(total) if total is not None else 0
            percentage = int(percentage) if percentage is not None else 0

            # 添加调试日志（仅在异常情况下记录，避免日志过多）
            if percentage > 100 or percentage < -1:
                logger.warning(f"异常的进度数据 - Downloaded: {downloaded}, Total: {total}, Percentage: {percentage}")

        except (ValueError, TypeError) as e:
            logger.warning(f"进度数据类型异常: {e} - 原始数据: downloaded={downloaded}, total={total}, percentage={percentage}")
            downloaded = total = percentage = 0

        if percentage >= 0:  # 正常进度更新
            # 限制百分比在合理范围内
            safe_percentage = max(0, min(100, percentage))

            # 重置进度条为正常模式
            if self.progress_bar.minimum() != 0 or self.progress_bar.maximum() != 100:
                self.progress_bar.setRange(0, 100)

            self.progress_bar.setValue(safe_percentage)

            # 格式化显示下载大小
            try:
                downloaded_mb = downloaded / (1024 * 1024)
                total_mb = total / (1024 * 1024)

                # 确保数值合理
                if downloaded_mb < 0 or total_mb < 0:
                    downloaded_mb = total_mb = 0

                # 如果总大小为0或无效，只显示已下载大小
                if total_mb <= 0:
                    progress_text = f"{downloaded_mb:.1f} MB (计算进度中...)"
                else:
                    progress_text = f"{downloaded_mb:.1f} MB / {total_mb:.1f} MB ({safe_percentage}%)"

                self.progress_label.setText(progress_text)
            except (ValueError, TypeError, ZeroDivisionError):
                self.progress_label.setText(f"进度: {safe_percentage}%")

        else:  # 等待状态（负数百分比表示等待状态）
            self.progress_bar.setRange(0, 0)  # 无限进度条
            # 使用通用等待消息，具体等待时间在下载管理器中处理
            self.progress_label.setText(f"{UpdateUIText.RETRYING_MESSAGE} 请稍候...")

    def _update_status(self, status: str) -> None:
        """
        更新状态

        Args:
            status: 状态信息
        """
        self.status_label.setText(status)
        logger.debug(f"更新状态: {status}")

    def _update_finished(self, success: bool, error: str, download_path: str) -> None:
        """
        更新完成处理

        Args:
            success: 是否成功
            error: 错误信息
            download_path: 下载文件路径
        """
        # 保存下载路径用于重启
        self.download_path = download_path

        if success:
            self.status_label.setText(UpdateUIText.UPDATE_COMPLETE_MESSAGE)
            self.progress_bar.setValue(100)
            self.progress_label.setText("100%")

            # 获取更新后的程序路径
            self.updated_executable_path = self._get_updated_executable_path()
            logger.info(f"更新完成，程序路径: {self.updated_executable_path}")

            # 显示下载路径信息
            if download_path:
                info_message = f"{UpdateUIText.FILE_DOWNLOADED_TO_MESSAGE}\n{download_path}"
                QMessageBox.information(self, UpdateUIText.UPDATE_COMPLETE_TITLE, info_message)

            # 2秒后关闭对话框并重启应用
            QTimer.singleShot(2000, self._restart_application)
        else:
            self.status_label.setText(f"{UpdateUIText.UPDATE_FAILED_MESSAGE}: {error}")
            self.progress_bar.setValue(0)
            self.progress_label.setText("")
            self.cancel_btn.setEnabled(True)
            QMessageBox.critical(self, UpdateUIText.UPDATE_FAILED_TITLE, f"{UpdateUIText.UPDATE_FAILED_MESSAGE}:\n{error}")

    def _get_updated_executable_path(self) -> str:
        """
        获取更新后的可执行文件路径

        Returns:
            更新后的程序文件路径

        逻辑说明:
        - 优先级1: 从待更新标记读取延迟更新路径（持久化，重启后可用）
        - 优先级2: 从类变量读取延迟更新路径（当前进程）
        - 优先级3: 生产环境的exe完整路径
        - 优先级4: 开发环境的exe或py文件
        """
        try:
            is_prod = is_production_environment()

            logger.info(f"[路径查找] 环境: {'生产环境' if is_prod else '开发环境'}")
            logger.info(f"[路径查找] 配置的exe文件名: {APP_EXECUTABLE}")
            logger.info(f"[路径查找] 配置的应用名: {APP_NAME}")

            # ✅ 优先级1: 从待更新标记文件读取延迟更新路径
            try:
                from .two_phase_updater import TwoPhaseUpdater
                two_phase = TwoPhaseUpdater()

                if two_phase.has_pending_update():
                    pending_info = two_phase.get_pending_update_info()
                    if pending_info and "source_file" in pending_info:
                        delayed_path = pending_info["source_file"]
                        if os.path.exists(delayed_path):
                            logger.info(f"[路径查找] ✓ 从标记文件读取延迟更新路径")
                            logger.info(f"[路径查找] 延迟更新路径: {delayed_path}")
                            logger.info(f"[路径查找] 将启动下载目录中的新版本")
                            return delayed_path
            except Exception as e:
                logger.warning(f"[路径查找] 从标记文件读取失败: {e}")

            # ✅ 优先级2: 检查类变量中的延迟更新路径
            if self.auto_updater and hasattr(self.auto_updater, 'update_executor'):
                delayed_path = self.auto_updater.update_executor.delayed_update_path
                if delayed_path and os.path.exists(delayed_path):
                    logger.info(f"[路径查找] ✓ 发现延迟更新路径（类变量）")
                    logger.info(f"[路径查找] 延迟更新路径: {delayed_path}")
                    logger.info(f"[路径查找] 将启动下载目录中的新版本")
                    return delayed_path

            if is_prod:
                # 生产环境:直接使用配置函数返回的完整路径
                # 不要使用 dirname 再拼接,这样会导致路径错误
                exe_path = get_app_executable_path()
                exec_dir = os.path.dirname(exe_path)

                logger.info(f"[生产环境] 完整exe路径: {exe_path}")
                logger.info(f"[生产环境] 所在目录: {exec_dir}")

                if os.path.exists(exe_path):
                    logger.info(f"[生产环境] ✓ exe文件存在")
                    return exe_path
                else:
                    logger.warning(f"[生产环境] ✗ exe文件不存在,尝试目录查找")
                    # 如果完整路径不存在,尝试在所在目录查找
                    alt_path = os.path.join(exec_dir, APP_EXECUTABLE)
                    if os.path.exists(alt_path):
                        logger.info(f"[生产环境] ✓ 在目录中找到exe: {alt_path}")
                        return alt_path
                    else:
                        logger.error(f"[生产环境] ✗ 未找到exe文件")
                        return exe_path  # 返回原路径,让启动逻辑报告错误
            else:
                # 开发环境:优先查找exe文件(用于测试更新功能),否则使用py文件
                exec_dir = os.path.dirname(get_app_executable_path())

                logger.info(f"[开发环境] 执行目录: {exec_dir}")

                # 优先查找exe文件(用于测试更新功能)
                exe_path = os.path.join(exec_dir, APP_EXECUTABLE)
                logger.info(f"[开发环境] 检查exe文件: {exe_path}")

                if os.path.exists(exe_path):
                    logger.info(f"[开发环境] ✓ 发现exe文件,将启动exe进行测试")
                    return exe_path

                # 没有exe文件,使用py文件
                py_path = os.path.join(exec_dir, f"{APP_NAME}.py")
                logger.info(f"[开发环境] exe文件不存在,检查py文件: {py_path}")

                if os.path.exists(py_path):
                    logger.info(f"[开发环境] ✓ 使用py文件: {py_path}")
                    return py_path
                else:
                    logger.warning(f"[开发环境] ✗ py文件不存在,回退到默认路径")
                    return get_app_executable_path()

        except Exception as e:
            logger.error(f"[路径查找] 异常: {e}", exc_info=True)
            logger.warning(f"[路径查找] 回退到默认路径")
            return get_app_executable_path()

    def _restart_application(self) -> None:
        """
        重启应用程序

        启动流程:
        1. 获取程序路径
        2. 验证路径有效性
        3. 根据文件类型启动
        4. 退出当前应用
        """
        try:
            logger.info("=" * 60)
            logger.info("[重启应用] 开始重启流程")
            logger.info("=" * 60)

            # 关闭对话框
            self.accept()

            # 获取要启动的程序路径
            executable_path = self._get_updated_executable_path()
            logger.info(f"[重启应用] 目标路径: {executable_path}")

            # 路径规范化
            executable_path = os.path.normpath(executable_path)
            logger.info(f"[重启应用] 规范化路径: {executable_path}")

            # 验证1: 文件是否存在
            if not os.path.exists(executable_path):
                error_msg = f"程序文件不存在:\n{executable_path}"
                logger.error(f"[重启应用] ✗ {error_msg}")
                QMessageBox.critical(
                    self,
                    UpdateUIText.RESTART_FAILED_TITLE,
                    error_msg
                )
                return

            # 验证2: 文件大小是否有效(至少1KB)
            file_size = os.path.getsize(executable_path)
            logger.info(f"[重启应用] 文件大小: {file_size} bytes ({file_size/1024:.2f} KB)")

            if file_size < 1024:
                error_msg = f"程序文件无效(文件过小):\n{executable_path}\n大小: {file_size} bytes"
                logger.error(f"[重启应用] ✗ {error_msg}")
                QMessageBox.critical(
                    self,
                    UpdateUIText.RESTART_FAILED_TITLE,
                    error_msg
                )
                return

            # 验证3: 工作目录是否存在
            working_dir = os.path.dirname(executable_path)
            if not os.path.exists(working_dir):
                error_msg = f"工作目录不存在:\n{working_dir}"
                logger.error(f"[重启应用] ✗ {error_msg}")
                QMessageBox.critical(
                    self,
                    UpdateUIText.RESTART_FAILED_TITLE,
                    error_msg
                )
                return

            logger.info(f"[重启应用] ✓ 文件验证通过")
            logger.info(f"[重启应用] 工作目录: {working_dir}")

            # 启动参数准备
            startup_info = {
                'env': os.environ.copy(),
                'cwd': working_dir
            }
            logger.info(f"[重启应用] 环境变量已复制")
            logger.info(f"[重启应用] 工作目录已设置")

            # 根据文件类型启动
            if executable_path.endswith('.exe'):
                # 启动exe文件 - Windows环境使用特殊处理
                logger.info(f"[重启应用] 启动类型: EXE文件")
                logger.info(f"[重启应用] 启动命令: {executable_path}")

                try:
                    # Windows下使用shell=True启动exe更可靠
                    # 这样可以正确处理路径中的空格和特殊字符
                    if sys.platform == 'win32':
                        logger.info(f"[重启应用] 使用Windows shell启动")
                        # 使用shell启动,路径会自动处理空格
                        process = subprocess.Popen(
                            f'"{executable_path}"',  # 路径加引号处理空格
                            shell=True,
                            cwd=working_dir,
                            env=os.environ.copy()
                        )
                    else:
                        # 非Windows环境使用常规方式
                        process = subprocess.Popen(
                            [executable_path],
                            cwd=working_dir,
                            env=os.environ.copy()
                        )

                    logger.info(f"[重启应用] ✓ 进程已启动, PID: {process.pid}")

                except FileNotFoundError:
                    # 如果shell启动失败,尝试使用os.startfile
                    logger.warning(f"[重启应用] shell启动失败,尝试os.startfile")
                    try:
                        import ctypes
                        ctypes.windll.shell32.ShellExecuteW(None, "open", executable_path, None, working_dir, 1)
                        logger.info(f"[重启应用] ✓ 通过ShellExecuteW启动")
                        # 使用ShellExecuteW无法获取PID,所以跳过
                    except Exception as shell_error:
                        logger.error(f"[重启应用] ✗ ShellExecuteW启动失败: {shell_error}")
                        raise

            elif executable_path.endswith('.py'):
                # 启动Python文件
                logger.info(f"[重启应用] 启动类型: Python文件")
                logger.info(f"[重启应用] Python解释器: {sys.executable}")
                logger.info(f"[重启应用] 启动命令: {sys.executable} {executable_path}")

                process = subprocess.Popen(
                    [sys.executable, executable_path],
                    cwd=working_dir,
                    env=os.environ.copy()
                )

                logger.info(f"[重启应用] ✓ 进程已启动, PID: {process.pid}")

            else:
                # 其他类型文件,直接启动
                logger.warning(f"[重启应用] 启动类型: 未知文件类型")
                logger.info(f"[重启应用] 启动命令: {executable_path}")

                process = subprocess.Popen(
                    [executable_path],
                    cwd=working_dir,
                    env=os.environ.copy()
                )

                logger.info(f"[重启应用] ✓ 进程已启动, PID: {process.pid}")

            logger.info("=" * 60)
            logger.info("[重启应用] 启动成功,准备退出当前应用")
            logger.info("=" * 60)

            # 延迟100ms后退出当前应用(确保新进程已启动)
            QTimer.singleShot(100, QApplication.quit)

        except Exception as e:
            logger.error(f"[重启应用] ✗ 重启失败: {e}", exc_info=True)
            QMessageBox.critical(
                self,
                UpdateUIText.RESTART_FAILED_TITLE,
                f"{UpdateUIText.RESTART_FAILED_MESSAGE}:\n{str(e)}"
            )
            self.reject()

    def cancel_update(self) -> None:
        """取消更新"""
        if self.update_thread and self.update_thread.isRunning():
            self.update_thread.cancel()
            self.update_thread.wait(3000)  # 等待3秒

        self.reject()


class AboutDialog(QDialog):
    """
    关于对话框

    显示应用程序版本信息和更新功能入口
    """

    def __init__(self, auto_updater, parent=None):
        """
        初始化关于对话框

        Args:
            auto_updater: AutoUpdater实例
            parent: 父窗口
        """
        super().__init__(parent)
        self.setWindowTitle(UpdateUIText.ABOUT_DIALOG_TITLE)
        self.setFixedSize(UpdateUIStyle.ABOUT_DIALOG_SIZE)
        self.setModal(True)
        self.auto_updater = auto_updater
        self.parent = parent

        self._setup_ui()
        self._load_version_info()

    def _setup_ui(self) -> None:
        """设置UI界面"""
        layout = QVBoxLayout()

        # 应用程序名称和版本
        title_label = QLabel(UpdateUIText.APP_NAME)
        title_label.setStyleSheet(UpdateUIStyle.TITLE_LABEL_STYLE)
        title_label.setAlignment(Qt.AlignCenter)

        # 版本信息
        self.version_label = QLabel(UpdateUIText.LOADING_VERSION_MESSAGE)
        self.version_label.setAlignment(Qt.AlignCenter)

        # 构建信息
        self.build_info_label = QLabel("")
        self.build_info_label.setAlignment(Qt.AlignCenter)
        self.build_info_label.setStyleSheet(UpdateUIStyle.BUILD_INFO_STYLE)

        # 更新状态
        self.update_status_label = QLabel(UpdateUIText.CHECKING_UPDATE_STATUS_MESSAGE)
        self.update_status_label.setAlignment(Qt.AlignCenter)

        # 分隔线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)

        # 详细信息区域
        info_layout = QVBoxLayout()

        # GitHub链接
        self.github_link_label = QLabel(UpdateUIText.GITHUB_LINK_TEXT)
        self.github_link_label.setStyleSheet(UpdateUIStyle.LINK_STYLE)
        self.github_link_label.setCursor(Qt.PointingHandCursor)

        # 最后检查时间
        self.last_check_label = QLabel("")

        # 配置信息
        self.config_info_label = QLabel("")

        info_layout.addWidget(self.github_link_label)
        info_layout.addWidget(self.last_check_label)
        info_layout.addWidget(self.config_info_label)

        # 按钮区域
        button_layout = QHBoxLayout()

        self.check_update_btn = QPushButton(UpdateUIText.CHECK_UPDATE_BUTTON_TEXT)
        self.check_update_btn.clicked.connect(self._check_for_updates)

        self.view_release_notes_btn = QPushButton(UpdateUIText.VIEW_RELEASE_NOTES_BUTTON_TEXT)
        self.view_release_notes_btn.clicked.connect(self._view_release_notes)

        self.close_btn = QPushButton(UpdateUIText.CLOSE_BUTTON_TEXT)
        self.close_btn.clicked.connect(self.accept)

        button_layout.addWidget(self.check_update_btn)
        button_layout.addWidget(self.view_release_notes_btn)
        button_layout.addWidget(self.close_btn)

        # 添加所有控件到布局
        layout.addWidget(title_label)
        layout.addWidget(self.version_label)
        layout.addWidget(self.build_info_label)
        layout.addWidget(self.update_status_label)
        layout.addWidget(line)
        layout.addLayout(info_layout)
        layout.addStretch()
        layout.addLayout(button_layout)

        self.setLayout(layout)

        # 设置事件
        self.github_link_label.mousePressEvent = self._open_github_page

    def _load_version_info(self) -> None:
        """加载版本信息"""
        try:
            # 显示当前版本
            if self.auto_updater:
                local_version = self.auto_updater.config.current_version
                self.version_label.setText(f"{UpdateUIText.VERSION_PREFIX}{local_version}")

                # 异步检查更新状态
                self._check_update_status_async()
            else:
                from .. import get_version
                self.version_label.setText(f"{UpdateUIText.VERSION_PREFIX}{get_version()}")
                self.update_status_label.setText(UpdateUIText.UPDATE_UNAVAILABLE_MESSAGE)
                self.check_update_btn.setEnabled(False)

        except Exception as e:
            from .. import get_version
            self.version_label.setText(f"{UpdateUIText.VERSION_PREFIX}{get_version()}")
            self.update_status_label.setText(f"{UpdateUIText.GET_VERSION_ERROR_MESSAGE}: {str(e)}")

    def _check_update_status_async(self) -> None:
        """异步检查更新状态"""
        try:
            # 使用定时器延迟执行，避免阻塞UI
            QTimer.singleShot(100, self._perform_status_check)
        except Exception as e:
            self.update_status_label.setText(f"{UpdateUIText.CHECK_UPDATE_STATUS_ERROR_MESSAGE}: {str(e)}")

    def _perform_status_check(self) -> None:
        """
        执行更新状态检查（手动检查）

        这是关于对话框中的手动检查，需要强制检查更新
        """
        try:
            if not self.auto_updater:
                return

            # 手动检查：强制检查，无视间隔限制
            has_update, remote_version, local_version, error = self.auto_updater.check_for_updates(force_check=True)

            if error:
                self.update_status_label.setText(UpdateUIText.CANNOT_CHECK_UPDATE_STATUS_MESSAGE)
                self.update_status_label.setStyleSheet(UpdateUIStyle.STATUS_ERROR_STYLE)
            elif has_update:
                self.update_status_label.setText(f"{UpdateUIText.NEW_VERSION_FOUND_MESSAGE_SIMPLE} {remote_version}")
                self.update_status_label.setStyleSheet(UpdateUIStyle.STATUS_UPDATE_STYLE)
            else:
                self.update_status_label.setText(UpdateUIText.LATEST_VERSION_MESSAGE_SIMPLE)
                self.update_status_label.setStyleSheet(UpdateUIStyle.STATUS_CURRENT_STYLE)

        except Exception as e:
            self.update_status_label.setText(f"{UpdateUIText.CHECK_UPDATE_STATUS_ERROR_MESSAGE}: {str(e)}")

    def _check_for_updates(self) -> None:
        """检查更新"""
        try:
            if self.parent and hasattr(self.parent, 'check_for_updates'):
                self.parent.check_for_updates()
                self.accept()  # 关闭关于对话框
            else:
                QMessageBox.warning(self, UpdateUIText.ERROR_TITLE, UpdateUIText.CHECK_UPDATE_FAILED_MESSAGE)
        except Exception as e:
            QMessageBox.warning(self, UpdateUIText.ERROR_TITLE, f"{UpdateUIText.CHECK_UPDATE_FAILED_MESSAGE}: {str(e)}")

    def _open_github_page(self, event) -> None:
        """打开GitHub页面"""
        try:
            webbrowser.open(UpdateUIText.GITHUB_URL)
        except Exception as e:
            QMessageBox.warning(self, UpdateUIText.ERROR_TITLE, f"{UpdateUIText.OPEN_GITHUB_ERROR_MESSAGE}: {str(e)}")

    def _view_release_notes(self) -> None:
        """
        查看更新日志（手动检查）

        这是用户主动点击查看更新日志，需要强制检查更新
        """
        try:
            if self.auto_updater:
                # 手动检查：强制检查，无视间隔限制
                has_update, remote_version, local_version, error = self.auto_updater.check_for_updates(force_check=True)

                if error:
                    QMessageBox.warning(self, UpdateUIText.ERROR_TITLE, f"{UpdateUIText.GET_RELEASE_NOTES_ERROR_MESSAGE}: {error}")
                    return

                # 获取最新版本的发布说明
                release_notes = self.auto_updater.github_client.get_release_notes(local_version)

                # 创建发布说明对话框
                dialog = QMessageBox(self)
                dialog.setWindowTitle(UpdateUIText.RELEASE_NOTES_TITLE)
                dialog.setText(f"{UpdateUIText.RELEASE_NOTES_MESSAGE} {local_version}:")
                dialog.setInformativeText(release_notes)
                dialog.setStandardButtons(QMessageBox.Ok)
                dialog.exec_()
            else:
                QMessageBox.information(self, UpdateUIText.INFO_TITLE, UpdateUIText.GET_RELEASE_NOTES_UNAVAILABLE_MESSAGE)

        except Exception as e:
            QMessageBox.warning(self, UpdateUIText.ERROR_TITLE, f"{UpdateUIText.VIEW_RELEASE_NOTES_ERROR_MESSAGE}: {str(e)}")