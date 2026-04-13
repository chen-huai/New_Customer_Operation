"""
Update UI Manager

Central UI controller for update operations. Manages all user interface
interactions during the update process including dialogs, progress displays,
and user notifications.
"""

from typing import Optional, Callable, Any
from PyQt5.QtWidgets import QMessageBox, QWidget, QProgressDialog
from PyQt5.QtCore import QTimer, QObject, Qt
import logging
from functools import wraps

from .update_dialogs import UpdateDialogs
from .progress_dialog import ProgressDialog

logger = logging.getLogger(__name__)


def safe_dialog_operation(func):
    """
    装饰器：安全地处理对话框操作，捕获RuntimeError
    """
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        try:
            return func(self, *args, **kwargs)
        except RuntimeError as e:
            if "wrapped C/C++ object has been deleted" in str(e):
                logger.info(f"对话框操作时对象已删除: {func.__name__}")
                return None
            else:
                logger.error(f"对话框操作发生RuntimeError: {func.__name__}: {e}")
                raise
        except Exception as e:
            logger.error(f"对话框操作发生意外错误: {func.__name__}: {e}")
            return None
    return wrapper


class UpdateUIManager:
    """UI管理器，负责更新过程中的所有用户交互"""

    def __init__(self, parent: Optional[QWidget] = None):
        """
        初始化UI管理器

        Args:
            parent: 父窗口对象，用于居中显示对话框
        """
        self.parent = parent
        self.update_dialogs = UpdateDialogs(parent)
        self.progress_dialog = ProgressDialog(parent)

    def set_parent(self, parent: QWidget):
        """设置父窗口"""
        self.parent = parent
        self.update_dialogs.set_parent(parent)
        self.progress_dialog.set_parent(parent)

    def show_update_notification(self, has_update: bool, remote_version: str = None,
                                error_msg: str = None) -> tuple:
        """
        显示更新通知

        Args:
            has_update: 是否有更新
            remote_version: 远程版本号
            error_msg: 错误信息

        Returns:
            tuple: (用户选择结果, 错误信息)
        """
        return self.update_dialogs.show_update_notification(has_update, remote_version, error_msg)

    def show_download_confirm(self, version: str, file_size: str = None, show_loading: bool = False) -> bool:
        """
        显示下载确认对话框

        Args:
            version: 版本号
            file_size: 文件大小或加载信息
            show_loading: 是否显示加载状态

        Returns:
            bool: 用户是否确认下载
        """
        return self.update_dialogs.show_download_confirm(version, file_size, show_loading)

    def show_install_confirm(self, version: str) -> bool:
        """
        显示安装确认对话框

        Args:
            version: 版本号

        Returns:
            bool: 用户是否确认安装
        """
        return self.update_dialogs.show_install_confirm(version)

    def show_update_complete(self, version: str, needs_restart: bool = True) -> None:
        """
        显示更新完成对话框

        Args:
            version: 更新到的版本
            needs_restart: 是否需要重启应用
        """
        self.update_dialogs.show_update_complete(version, needs_restart)

    def create_progress_dialog(self, title: str = "下载更新") -> Any:
        """
        创建进度对话框

        Args:
            title: 对话框标题

        Returns:
            进度对话框对象
        """
        return self.progress_dialog.create_progress_dialog(title)

    def update_progress(self, dialog: Any, downloaded: int, total: int,
                       percentage: float, extra_info: str = None) -> None:
        """
        更新进度显示

        Args:
            dialog: 进度对话框对象
            downloaded: 已下载字节数
            total: 总字节数
            percentage: 完成百分比
            extra_info: 额外信息
        """
        self.progress_dialog.update_progress(dialog, downloaded, total, percentage, extra_info)

    def close_progress_dialog(self, dialog: Any) -> None:
        """
        关闭进度对话框

        Args:
            dialog: 进度对话框对象
        """
        self.progress_dialog.close_progress_dialog(dialog)

    def show_error_dialog(self, title: str, message: str, error_type: str = "error") -> None:
        """
        显示错误对话框

        Args:
            title: 对话框标题
            message: 错误信息
            error_type: 错误类型 (warning/error/critical)
        """
        self.update_dialogs.show_error_dialog(title, message, error_type)

    def format_file_size(self, size_bytes: int) -> str:
        """
        格式化文件大小显示

        Args:
            size_bytes: 文件大小（字节）

        Returns:
            str: 格式化后的文件大小
        """
        return self.progress_dialog.format_file_size(size_bytes)

    def download_with_dialog_async(self, download_thread, version: str,
                                  progress_callback: Optional[Callable] = None, max_retries: int = 3) -> tuple:
        """
        使用异步对话框进行下载

        Args:
            download_thread: 异步下载线程对象
            version: 版本号
            progress_callback: 进度回调函数（可选）
            max_retries: 最大重试次数

        Returns:
            tuple: (下载是否成功, 下载的文件路径, 错误信息)
        """
        logger.info(f"开始异步下载流程，版本: {version}")

        try:
            # 创建异步进度对话框
            progress_dialog = self.progress_dialog.create_async_progress_dialog("下载更新", allow_cancel=True)

            # 存储下载结果
            download_result = {'success': False, 'file_path': '', 'error': None}
            current_thread = [download_thread]  # 使用列表以便在闭包中修改

            def on_progress_updated(downloaded: int, total: int, percentage: float):
                """进度更新回调（修复版本）"""
                if progress_dialog is None:
                    return

                try:
                    # 安全检查对话框状态
                    if progress_dialog.wasCanceled():
                        # 用户取消了下载
                        if current_thread[0] and hasattr(current_thread[0], 'cancel_download'):
                            current_thread[0].cancel_download()
                        return

                    # 更新进度
                    if not progress_dialog.wasCanceled():
                        self.progress_dialog.update_async_progress(progress_dialog, downloaded, total, percentage)

                    # 调用外部回调
                    if progress_callback:
                        progress_callback(downloaded, total, percentage)

                except RuntimeError as e:
                    if "wrapped C/C++ object has been deleted" in str(e):
                        logger.info("进度对话框已删除，停止进度更新")
                        return
                    else:
                        logger.error(f"进度更新时发生RuntimeError: {e}")
                        return
                except Exception as e:
                    logger.error(f"进度更新时发生意外错误: {e}")
                    return

            def on_speed_calculated(speed_mbs: float):
                """下载速度回调"""
                # 速度信息会在progress对话框中自动更新
                pass

            @safe_dialog_operation
            def on_retry_attempt(current_retry: int, max_retries: int):
                """重试尝试回调（简化版本）"""
                if progress_dialog:
                    status_msg = f"网络重试 {current_retry}/{max_retries}"
                    self.progress_dialog.update_status_async(progress_dialog, status_msg)

            @safe_dialog_operation
            def on_status_changed(status: str):
                """状态变更回调（简化版本）"""
                if progress_dialog and not progress_dialog.wasCanceled():
                    self.progress_dialog.update_status_async(progress_dialog, status)

            def on_download_complete(success: bool, file_path: str, error: str):
                """下载完成回调（修复版本）"""
                logger.info(f"下载完成: success={success}, file_path={file_path}, error={error}")
                download_result['success'] = success
                download_result['file_path'] = file_path
                download_result['error'] = error

                # 延迟关闭对话框，让用户看到完成状态
                if progress_dialog:
                    try:
                        if success:
                            self.progress_dialog.update_status_async(progress_dialog, "下载完成！")
                            QTimer.singleShot(1000, lambda: self.progress_dialog.close_async_progress_dialog(progress_dialog))
                        else:
                            self.progress_dialog.close_async_progress_dialog(progress_dialog)
                    except RuntimeError as e:
                        if "wrapped C/C++ object has been deleted" in str(e):
                            logger.info("下载完成处理时对话框已删除")
                        else:
                            logger.error(f"下载完成处理时发生RuntimeError: {e}")
                    except Exception as e:
                        logger.error(f"下载完成处理时发生意外错误: {e}")

            def on_error_occurred(error_msg: str):
                """错误发生回调（修复版本）"""
                logger.error(f"下载错误: {error_msg}")
                download_result['success'] = False
                download_result['error'] = error_msg

                if progress_dialog:
                    try:
                        self.progress_dialog.update_status_async(progress_dialog, f"下载失败: {error_msg}")
                        QTimer.singleShot(2000, lambda: self.progress_dialog.close_async_progress_dialog(progress_dialog))
                    except RuntimeError as e:
                        if "wrapped C/C++ object has been deleted" in str(e):
                            logger.info("错误处理时对话框已删除")
                        else:
                            logger.error(f"错误处理时发生RuntimeError: {e}")
                    except Exception as e:
                        logger.error(f"错误处理时发生意外错误: {e}")

            # 使用Qt.QueuedConnection确保跨线程安全
            connection_type = Qt.QueuedConnection

            # 连接信号（确保线程安全）
            if hasattr(download_thread, 'progress_updated'):
                download_thread.progress_updated.connect(on_progress_updated, type=connection_type)
            if hasattr(download_thread, 'speed_calculated'):
                download_thread.speed_calculated.connect(on_speed_calculated, type=connection_type)
            if hasattr(download_thread, 'retry_attempt'):
                download_thread.retry_attempt.connect(on_retry_attempt, type=connection_type)
            if hasattr(download_thread, 'status_changed'):
                download_thread.status_changed.connect(on_status_changed, type=connection_type)
            if hasattr(download_thread, 'download_complete'):
                download_thread.download_complete.connect(on_download_complete, type=connection_type)
            if hasattr(download_thread, 'error_occurred'):
                download_thread.error_occurred.connect(on_error_occurred, type=connection_type)

            # 连接取消信号
            progress_dialog.canceled.connect(self._on_download_cancelled, type=connection_type)

            # 启动下载线程
            download_thread.start()

            # 等待下载完成（非阻塞方式）
            self._wait_for_download_completion(download_thread, progress_dialog, download_result)

            return download_result['success'], download_result['file_path'], download_result['error']

        except Exception as e:
            logger.error(f"异步下载流程异常: {e}", exc_info=True)
            return False, None, str(e)

    def _wait_for_download_completion(self, download_thread, progress_dialog: QProgressDialog,
                                    download_result: dict):
        """
        等待下载完成（非阻塞，修复版本）

        Args:
            download_thread: 下载线程
            progress_dialog: 进度对话框
            download_result: 下载结果字典
        """
        def check_completion():
            try:
                # 检查对话框是否仍然有效
                if progress_dialog is None:
                    logger.debug("进度对话框为空，停止检查")
                    return

                try:
                    # 安全地检查对话框状态
                    if progress_dialog.wasCanceled():
                        logger.info("用户取消了下载")
                        if download_thread and hasattr(download_thread, 'cancel_download'):
                            download_thread.cancel_download()
                        return
                except RuntimeError as e:
                    if "wrapped C/C++ object has been deleted" in str(e):
                        logger.info("进度对话框已被删除，停止检查")
                        return
                    else:
                        raise

                # 检查下载是否完成
                if download_thread and not download_thread.isRunning():
                    logger.info("下载线程已完成")
                    return

                # 如果还在运行，继续检查
                if download_thread and download_thread.isRunning():
                    QTimer.singleShot(100, check_completion)

            except Exception as e:
                logger.error(f"检查下载完成状态时出错: {e}")
                # 出错时也停止检查，避免无限循环

        # 使用QTimer确保在主线程中执行检查
        QTimer.singleShot(100, check_completion)

    def _on_download_cancelled(self):
        """下载取消处理"""
        logger.info("用户取消了下载")
        # 具体的取消逻辑由下载线程处理

    def show_download_status(self, title: str, message: str) -> QProgressDialog:
        """
        显示下载状态对话框

        Args:
            title: 对话框标题
            message: 状态信息

        Returns:
            QProgressDialog: 状态对话框对象
        """
        try:
            status_dialog = self.progress_dialog.create_async_progress_dialog(title, allow_cancel=False)
            self.progress_dialog.update_status_async(status_dialog, message)
            return status_dialog
        except Exception as e:
            logger.error(f"显示下载状态对话框失败: {e}")
            raise