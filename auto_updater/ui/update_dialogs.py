"""
Update Dialogs

Provides various dialog boxes for update operations including update notifications,
confirmations, error messages, and completion dialogs.
"""

from typing import Optional, Tuple
from PyQt5.QtWidgets import QMessageBox, QWidget
from PyQt5.QtCore import QTimer
import logging

logger = logging.getLogger(__name__)


class UpdateDialogs:
    """更新对话框管理器"""

    def __init__(self, parent: Optional[QWidget] = None):
        """
        初始化对话框管理器

        Args:
            parent: 父窗口对象
        """
        self.parent = parent

    def set_parent(self, parent: QWidget):
        """设置父窗口"""
        self.parent = parent

    def show_update_notification(self, has_update: bool, remote_version: str = None,
                                error_msg: str = None) -> Tuple[bool, str]:
        """
        显示更新通知

        Args:
            has_update: 是否有更新
            remote_version: 远程版本号
            error_msg: 错误信息

        Returns:
            tuple: (用户选择结果, 错误信息)
        """
        try:
            if error_msg and "距离上次检查时间过短" not in error_msg:
                if "网络" in error_msg or "连接" in error_msg:
                    self._show_network_error(error_msg)
                else:
                    self._show_general_error(error_msg)
                return False, error_msg

            if has_update and remote_version:
                return self._show_update_available_dialog(remote_version), None
            else:
                return self._show_no_update_dialog(), None

        except Exception as e:
            logger.error(f"显示更新通知失败: {e}")
            return False, str(e)

    def _show_update_available_dialog(self, version: str) -> bool:
        """显示有新版本对话框"""
        msg_box = QMessageBox(self.parent)
        msg_box.setIcon(QMessageBox.Question)
        msg_box.setWindowTitle("发现新版本")
        msg_box.setText(f"发现新版本 {version}！")
        msg_box.setInformativeText("是否立即下载更新？")
        msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg_box.setDefaultButton(QMessageBox.Yes)

        return msg_box.exec_() == QMessageBox.Yes

    def _show_no_update_dialog(self) -> bool:
        """显示无更新对话框"""
        msg_box = QMessageBox(self.parent)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle("检查更新")
        msg_box.setText("您当前已是最新版本！")
        msg_box.setStandardButtons(QMessageBox.Ok)

        msg_box.exec_()
        return False

    def _show_network_error(self, error_msg: str):
        """显示网络错误"""
        msg_box = QMessageBox(self.parent)
        msg_box.setIcon(QMessageBox.Warning)
        msg_box.setWindowTitle("网络错误")
        msg_box.setText(f"网络连接出现问题：{error_msg}")
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    def _show_general_error(self, error_msg: str):
        """显示一般错误"""
        msg_box = QMessageBox(self.parent)
        msg_box.setIcon(QMessageBox.Warning)
        msg_box.setWindowTitle("更新检查失败")
        msg_box.setText(f"检查更新时发生错误：{error_msg}")
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    def _setup_yes_no_buttons(self, msg_box: QMessageBox) -> None:
        """
        设置标准的确认/取消按钮

        Args:
            msg_box: QMessageBox实例
        """
        msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg_box.setDefaultButton(QMessageBox.Yes)

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
        try:
            msg_box = QMessageBox(self.parent)

            if show_loading:
                msg_box.setIcon(QMessageBox.Information)
                msg_box.setWindowTitle("下载确认")
                msg_box.setText(f"准备下载版本 {version}")
                if file_size:
                    msg_box.setInformativeText(f"文件大小：{file_size}<br><br>是否继续下载？")
                else:
                    msg_box.setInformativeText("是否继续下载？")
            else:
                msg_box.setIcon(QMessageBox.Question)
                msg_box.setWindowTitle("确认下载")
                msg_box.setText(f"是否下载版本 {version}？")
                if file_size:
                    msg_box.setInformativeText(f"文件大小：{file_size}")

            # 统一设置按钮
            self._setup_yes_no_buttons(msg_box)

            return msg_box.exec_() == QMessageBox.Yes

        except Exception as e:
            logger.error(f"显示下载确认对话框失败: {e}")
            return False

    def show_install_confirm(self, version: str) -> bool:
        """
        显示安装确认对话框

        Args:
            version: 版本号

        Returns:
            bool: 用户是否确认安装
        """
        try:
            msg_box = QMessageBox(self.parent)
            msg_box.setIcon(QMessageBox.Question)
            msg_box.setWindowTitle("确认安装")
            msg_box.setText(f"下载完成，是否立即安装版本 {version}？")
            msg_box.setInformativeText("安装过程中应用程序会自动重启")
            msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            msg_box.setDefaultButton(QMessageBox.Yes)

            return msg_box.exec_() == QMessageBox.Yes

        except Exception as e:
            logger.error(f"显示安装确认对话框失败: {e}")
            return False

    def show_update_complete(self, version: str, needs_restart: bool = True) -> None:
        """
        显示更新完成对话框

        Args:
            version: 更新到的版本
            needs_restart: 是否需要重启应用
        """
        try:
            msg_box = QMessageBox(self.parent)
            msg_box.setIcon(QMessageBox.Information)
            msg_box.setWindowTitle("更新完成")

            if needs_restart:
                msg_box.setText(f"已成功更新到版本 {version}！")
                msg_box.setInformativeText("应用程序即将重启以完成更新")
            else:
                msg_box.setText(f"更新已完成！当前版本：{version}")

            msg_box.setStandardButtons(QMessageBox.Ok)
            msg_box.exec_()

        except Exception as e:
            logger.error(f"显示更新完成对话框失败: {e}")

    def show_error_dialog(self, title: str, message: str, error_type: str = "error") -> None:
        """
        显示错误对话框

        Args:
            title: 对话框标题
            message: 错误信息
            error_type: 错误类型 (warning/error/critical)
        """
        try:
            msg_box = QMessageBox(self.parent)

            # 设置图标
            if error_type == "warning":
                msg_box.setIcon(QMessageBox.Warning)
            else:
                msg_box.setIcon(QMessageBox.Critical)

            msg_box.setWindowTitle(title)
            msg_box.setText(message)
            msg_box.setStandardButtons(QMessageBox.Ok)
            msg_box.exec_()

        except Exception as e:
            logger.error(f"显示错误对话框失败: {e}")