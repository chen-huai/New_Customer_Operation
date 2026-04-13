"""
Progress Dialog

Provides progress display functionality for download operations including
file size formatting, progress percentage calculation, and UI updates.
"""

from typing import Optional
from PyQt5.QtWidgets import QMessageBox, QWidget, QProgressDialog
from PyQt5.QtCore import Qt
import logging

logger = logging.getLogger(__name__)


class ProgressDialog:
    """进度对话框管理器"""

    def __init__(self, parent: Optional[QWidget] = None):
        """
        初始化进度对话框管理器

        Args:
            parent: 父窗口对象
        """
        self.parent = parent
        self._last_update_time = 0
        self._update_interval = 0.1  # UI更新间隔（秒）- 动态调整
        self._min_update_interval = 0.05  # 最小更新间隔50ms
        self._max_update_interval = 0.5   # 最大更新间隔500ms
        self.download_speed = 0.0
        self.estimated_time_remaining = 0
        # 跟踪活跃的对话框实例，防止重复删除
        self._active_dialogs = set()
        # 性能优化：缓存UI更新频率
        self._performance_mode = False

    def set_parent(self, parent: QWidget):
        """设置父窗口"""
        self.parent = parent

    def create_progress_dialog(self, title: str = "下载更新") -> QMessageBox:
        """
        创建进度对话框

        Args:
            title: 对话框标题

        Returns:
            QMessageBox: 进度对话框对象
        """
        try:
            msg_box = QMessageBox(self.parent)
            msg_box.setIcon(QMessageBox.Information)
            msg_box.setWindowTitle(title)
            msg_box.setText("准备下载...")
            msg_box.setStandardButtons(QMessageBox.NoButton)
            msg_box.show()

            return msg_box
        except Exception as e:
            logger.error(f"创建进度对话框失败: {e}")
            raise

    def update_progress(self, dialog: QMessageBox, downloaded: int, total: int,
                       percentage: float, extra_info: str = None) -> None:
        """
        更新进度显示（优化版本）

        Args:
            dialog: 进度对话框
            downloaded: 已下载字节数
            total: 总字节数
            percentage: 完成百分比
            extra_info: 额外信息
        """
        try:
            import time

            current_time = time.time()
            # 性能优化：控制更新频率
            if current_time - self._last_update_time < self._update_interval:
                return

            self._last_update_time = current_time

            # 格式化文件大小和进度文本
            downloaded_str = self.format_file_size(downloaded)
            total_str = self.format_file_size(total)

            # 构建进度文本
            if extra_info:
                progress_text = f"{extra_info}\n{percentage:.1f}%\n已下载: {downloaded_str} / {total_str}"
            else:
                progress_text = f"正在下载更新... {percentage:.1f}%\n已下载: {downloaded_str} / {total_str}"

            # 更新对话框文本
            dialog.setText(progress_text)

            # 确保对话框显示在最前面
            dialog.raise_()
            dialog.activateWindow()

        except Exception as e:
            logger.error(f"更新进度显示失败: {e}")

    def close_progress_dialog(self, dialog: Optional[QMessageBox]) -> None:
        """
        关闭进度对话框

        Args:
            dialog: 进度对话框对象
        """
        try:
            if dialog and hasattr(dialog, 'close'):
                dialog.close()
        except Exception as e:
            logger.error(f"关闭进度对话框失败: {e}")

    def format_file_size(self, size_bytes: int) -> str:
        """
        格式化文件大小显示

        Args:
            size_bytes: 文件大小（字节）

        Returns:
            str: 格式化后的文件大小
        """
        try:
            if size_bytes == 0:
                return "0 B"

            # 定义单位
            units = ['B', 'KB', 'MB', 'GB', 'TB']
            unit_index = 0
            size = float(size_bytes)

            # 计算合适的单位
            while size >= 1024 and unit_index < len(units) - 1:
                size /= 1024
                unit_index += 1

            # 格式化输出
            if unit_index == 0:  # 字节
                return f"{int(size)} {units[unit_index]}"
            else:  # KB及以上
                return f"{size:.1f} {units[unit_index]}"

        except Exception as e:
            logger.error(f"格式化文件大小失败: {e}")
            return f"{size_bytes} B"

    def create_status_update_dialog(self, title: str, message: str) -> QMessageBox:
        """
        创建状态更新对话框

        Args:
            title: 对话框标题
            message: 显示信息

        Returns:
            QMessageBox: 状态对话框对象
        """
        try:
            msg_box = QMessageBox(self.parent)
            msg_box.setIcon(QMessageBox.Information)
            msg_box.setWindowTitle(title)
            msg_box.setText(message)
            msg_box.setStandardButtons(QMessageBox.NoButton)
            msg_box.show()

            return msg_box
        except Exception as e:
            logger.error(f"创建状态更新对话框失败: {e}")
            raise

    def update_status_message(self, dialog: QMessageBox, message: str) -> None:
        """
        更新状态信息

        Args:
            dialog: 对话框对象
            message: 新的状态信息
        """
        try:
            dialog.setText(message)
            dialog.raise_()
            dialog.activateWindow()
        except Exception as e:
            logger.error(f"更新状态信息失败: {e}")

    def create_async_progress_dialog(self, title: str = "下载更新", allow_cancel: bool = True) -> QProgressDialog:
        """
        创建异步下载进度对话框（支持取消功能）

        Args:
            title: 对话框标题
            allow_cancel: 是否允许取消

        Returns:
            QProgressDialog: 进度对话框对象
        """
        try:
            progress_dialog = QProgressDialog(self.parent)
            progress_dialog.setWindowTitle(title)
            progress_dialog.setLabelText("正在准备下载...")
            progress_dialog.setRange(0, 100)
            progress_dialog.setValue(0)

            if allow_cancel:
                progress_dialog.setCancelButtonText("取消")
                progress_dialog.canceled.connect(self._on_cancel_clicked)
            else:
                progress_dialog.setCancelButton(None)

            progress_dialog.setWindowModality(Qt.WindowModal)
            progress_dialog.setMinimumDuration(0)  # 立即显示
            progress_dialog.show()

            # 添加到活跃对话框跟踪
            self._active_dialogs.add(id(progress_dialog))
            logger.info(f"异步进度对话框创建成功: {title}, 活跃对话框数: {len(self._active_dialogs)}")
            return progress_dialog

        except Exception as e:
            logger.error(f"创建异步进度对话框失败: {e}")
            raise

    def update_async_progress(self, dialog: QProgressDialog, downloaded: int, total: int,
                            percentage: float, speed_mbs: float = 0.0) -> None:
        """
        更新异步进度显示（性能优化版本）

        Args:
            dialog: 进度对话框
            downloaded: 已下载字节数
            total: 总字节数
            percentage: 完成百分比
            speed_mbs: 下载速度 (MB/s)
        """
        try:
            import time

            current_time = time.time()

            # 自适应更新频率：高速度时减少更新频率
            if speed_mbs > 5.0:  # 高速下载时降低更新频率
                self._update_interval = min(self._max_update_interval, 0.2)
                self._performance_mode = True
            elif speed_mbs < 1.0:  # 低速下载时提高更新频率
                self._update_interval = max(self._min_update_interval, 0.08)
                self._performance_mode = False
            else:
                self._update_interval = 0.1  # 正常频率

            # 控制更新频率
            if current_time - self._last_update_time < self._update_interval:
                return

            self._last_update_time = current_time
            self.download_speed = speed_mbs

            # 更新进度条（只在数值变化时更新）
            current_value = dialog.value()
            new_value = int(percentage)
            if current_value != new_value:
                dialog.setValue(new_value)

            # 性能模式：简化显示信息
            if self._performance_mode:
                # 高速下载时只显示关键信息
                progress_text = f"正在下载更新... {percentage:.1f}% ({self.format_file_size(downloaded)})"
            else:
                # 正常模式：显示详细信息
                downloaded_str = self.format_file_size(downloaded)
                total_str = self.format_file_size(total)

                progress_lines = [f"正在下载更新... {percentage:.1f}%"]
                progress_lines.append(f"已下载: {downloaded_str} / {total_str}")

                if speed_mbs > 0:
                    progress_lines.append(f"下载速度: {speed_mbs:.2f} MB/s")

                    # 计算剩余时间（仅在低速时显示）
                    if speed_mbs < 2.0 and total > 0 and percentage < 100:
                        remaining_bytes = total - downloaded
                        remaining_seconds = remaining_bytes / (speed_mbs * 1024 * 1024)
                        remaining_str = self._format_time(remaining_seconds)
                        progress_lines.append(f"剩余时间: {remaining_str}")

                progress_text = "\n".join(progress_lines)

            dialog.setLabelText(progress_text)

            # 确保对话框显示在最前面
            dialog.raise_()
            dialog.activateWindow()

            # 性能优化：减少processEvents调用频率
            if not self._performance_mode or int(percentage) % 5 == 0:  # 高速时每5%更新一次
                from PyQt5.QtWidgets import QApplication
                app = QApplication.instance()
                if app:
                    app.processEvents()

        except Exception as e:
            logger.error(f"更新异步进度显示失败: {e}")

    def update_status_async(self, dialog: QProgressDialog, status: str) -> None:
        """
        更新异步状态信息

        Args:
            dialog: 进度对话框
            status: 状态信息
        """
        try:
            dialog.setLabelText(status)
            dialog.raise_()
            dialog.activateWindow()

            # 强制处理UI事件
            from PyQt5.QtWidgets import QApplication
            app = QApplication.instance()
            if app:
                app.processEvents()

        except Exception as e:
            logger.error(f"更新异步状态信息失败: {e}")

    def _is_dialog_valid(self, dialog: Optional[QProgressDialog]) -> bool:
        """
        检查对话框对象是否仍然有效

        Args:
            dialog: 对话框对象

        Returns:
            bool: 对话框是否有效
        """
        if dialog is None:
            return False

        # 检查对象是否已被删除
        try:
            dialog_id = id(dialog)
            if dialog_id not in self._active_dialogs:
                logger.warning(f"尝试关闭未跟踪的对话框: {dialog_id}")
                return False

            # 检查对象是否仍有效
            if hasattr(dialog, 'isVisible') and not dialog.isVisible():
                # 对象存在但不可见，可能已关闭
                self._active_dialogs.discard(dialog_id)
                return False

            return True
        except RuntimeError as e:
            if "wrapped C/C++ object has been deleted" in str(e):
                logger.info("对话框对象已被Qt删除，从跟踪中移除")
                self._active_dialogs.discard(id(dialog))
                return False
            else:
                logger.error(f"检查对话框有效性时出错: {e}")
                return False
        except Exception as e:
            logger.error(f"检查对话框有效性时发生意外错误: {e}")
            return False

    def close_async_progress_dialog(self, dialog: Optional[QProgressDialog]) -> None:
        """
        关闭异步进度对话框（修复版本）

        Args:
            dialog: 进度对话框对象
        """
        if dialog is None:
            return

        try:
            # 首先检查对话框是否仍然有效
            if not self._is_dialog_valid(dialog):
                logger.debug("对话框已无效，无需关闭")
                return

            dialog_id = id(dialog)

            # 安全关闭对话框
            if hasattr(dialog, 'close'):
                try:
                    dialog.close()
                    logger.debug(f"对话框关闭成功: {dialog_id}")
                except RuntimeError as e:
                    if "wrapped C/C++ object has been deleted" in str(e):
                        logger.info(f"关闭时发现对话框已被删除: {dialog_id}")
                    else:
                        raise

            # 延迟安全清理，避免在可能仍在使用时删除对象
            def safe_cleanup():
                try:
                    if dialog_id in self._active_dialogs:
                        self._active_dialogs.discard(dialog_id)
                        logger.debug(f"对话框从跟踪中移除: {dialog_id}")
                except Exception as e:
                    logger.error(f"清理对话框跟踪时出错: {e}")

            # 使用QTimer延迟执行清理，确保不在回调中直接删除
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(100, safe_cleanup)

        except Exception as e:
            logger.error(f"关闭异步进度对话框失败: {e}")
            # 确保从跟踪中移除，避免内存泄漏
            try:
                if dialog:
                    self._active_dialogs.discard(id(dialog))
            except:
                pass

    def _on_cancel_clicked(self):
        """取消按钮点击处理"""
        logger.info("用户点击了取消下载")
        # 这个方法会在UI线程中调用，具体处理由外部负责

    def _format_time(self, seconds: float) -> str:
        """
        格式化时间显示

        Args:
            seconds: 秒数

        Returns:
            str: 格式化后的时间字符串
        """
        try:
            if seconds < 60:
                return f"{int(seconds)}秒"
            elif seconds < 3600:
                minutes = int(seconds // 60)
                remaining_seconds = int(seconds % 60)
                return f"{minutes}分{remaining_seconds}秒"
            else:
                hours = int(seconds // 3600)
                remaining_minutes = int((seconds % 3600) // 60)
                return f"{hours}小时{remaining_minutes}分钟"
        except Exception:
            return "未知时间"