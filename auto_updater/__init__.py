# -*- coding: utf-8 -*-
"""
PDF重命名工具自动更新模块
提供基于GitHub Releases的自动更新功能
"""

from .config import get_config, is_development_environment
from .github_client import GitHubClient
from .download_manager import DownloadManager
from .backup_manager import BackupManager
from .update_executor import UpdateExecutor
from .config import *

# 自定义异常类
class UpdateError(Exception):
    """更新功能基础异常"""
    pass

class NetworkError(UpdateError):
    """网络连接异常"""
    pass

class VersionCheckError(UpdateError):
    """版本检查异常"""
    pass

class DownloadError(UpdateError):
    """文件下载异常"""
    pass

class BackupError(UpdateError):
    """备份操作异常"""
    pass

class UpdateExecutionError(UpdateError):
    """更新执行异常"""
    pass

# 主要接口类
class AutoUpdater:
    """
    自动更新器主类
    整合所有更新功能组件，提供统一的更新接口
    """

    def __init__(self, parent=None):
        """
        初始化自动更新器
        :param parent: 父对象（用于GUI信号连接）
        """
        self.config = get_config()
        self.github_client = GitHubClient()
        self.download_manager = DownloadManager()
        self.backup_manager = BackupManager()
        self.update_executor = UpdateExecutor()
        self.parent = parent

        # UI管理器（延迟初始化）
        self._ui_manager = None

    def check_for_updates(self, force_check=False, is_silent=False) -> tuple:
        """
        检查更新
        :param force_check: 是否强制检查（忽略时间间隔）
        :param is_silent: 是否静默模式（不返回间隔错误信息，用于启动检查）
        :return: (是否有更新, 远程版本, 本地版本, 错误信息)
        """
        try:
            local_version = self.config.current_version

            # 开发环境下的调试信息
            if is_development_environment():
                print(f"[开发环境] 开始检查更新，本地版本: {local_version}")
                if force_check:
                    print("[开发环境] 强制检查模式")
                if is_silent:
                    print("[开发环境] 静默模式")

            # 检查是否应该进行更新检查
            if not force_check:
                should_check, reason = self.config.should_check_for_updates()
                if not should_check:
                    if is_development_environment():
                        # 开发环境下不应该到达这里，因为should_check_for_updates总是返回True
                        # 这只是保险措施
                        print("[开发环境] 意外触发了时间间隔检查")
                        return False, None, local_version, "开发环境检查限制"
                    else:
                        # 静默模式下不返回间隔错误信息
                        if is_silent:
                            return False, None, local_version, None
                        else:
                            return False, None, local_version, f"距离上次检查时间过短（间隔{self.config.update_check_interval_days}天）"

            # 获取远程版本信息
            release_info = self.github_client.get_latest_release()
            if not release_info:
                return False, None, local_version, "无法获取远程版本信息"

            remote_version = release_info.get('tag_name', '').lstrip('v')
            if not remote_version:
                return False, None, local_version, "远程版本格式无效"

            # 检查是否有更新
            has_update = self.config.is_newer_version(remote_version, local_version)

            # 更新最后检查时间
            self.config.update_last_check_time()

            return has_update, remote_version, local_version, None

        except Exception as e:
            local_version = self.config.current_version
            return False, None, local_version, f"检查更新失败: {str(e)}"

    def download_update(self, version: str, progress_callback=None) -> tuple:
        """
        下载更新文件
        :param version: 要下载的版本号
        :param progress_callback: 进度回调函数
        :return: (是否成功, 下载文件路径, 错误信息)
        """
        try:
            # 获取下载链接
            download_url = self.github_client.get_download_url(version)
            if not download_url:
                return False, None, "无法获取下载链接"

            # 创建备份
            backup_path = self.backup_manager.create_backup()
            if not backup_path:
                return False, None, "创建备份失败"

            # 下载文件
            downloaded_file = self.download_manager.download_file(
                download_url,
                version,
                progress_callback
            )

            if downloaded_file:
                return True, downloaded_file, None
            else:
                return False, None, "下载失败"

        except Exception as e:
            return False, None, f"下载更新失败: {str(e)}"

    def execute_update(self, update_file_path: str, new_version: str) -> tuple:
        """
        执行应用程序更新操作

        Args:
            update_file_path (str): 下载的更新文件完整路径
            new_version (str): 目标版本号，必须符合语义化版本格式 (如 "1.2.3")

        Returns:
            tuple[bool, Optional[str]]: (更新是否成功, 错误信息)

        Raises:
            ValueError: 当版本号格式无效时
            UpdateExecutionError: 当更新执行过程中出现错误时

        Note:
            此方法会自动创建备份并执行文件替换操作
            更新成功后会更新本地版本信息
        """
        try:
            # 参数验证
            if not update_file_path or not update_file_path.strip():
                return False, "更新文件路径不能为空"

            if not new_version or not new_version.strip():
                return False, "新版本号不能为空"

            if not self._is_valid_version_format(new_version):
                return False, f"版本号格式无效: {new_version}"

            # 检查更新文件是否存在
            import os
            if not os.path.exists(update_file_path):
                return False, f"更新文件不存在: {update_file_path}"

            # 执行更新
            success = self.update_executor.execute_update(update_file_path, new_version)
            if success:
                return True, None
            else:
                return False, "更新执行失败"

        except UpdateExecutionError as e:
            # 保留具体的执行错误信息
            return False, f"更新执行失败: {str(e)}"
        except ValueError as e:
            # 参数验证错误处理
            return False, f"参数验证失败: {str(e)}"
        except Exception as e:
            return False, f"执行更新异常: {str(e)}"

    def rollback_update(self) -> tuple:
        """
        回滚更新
        :return: (是否成功, 错误信息)
        """
        try:
            success = self.backup_manager.restore_from_backup()
            if success:
                return True, None
            else:
                return False, "回滚失败"

        except Exception as e:
            return False, f"回滚异常: {str(e)}"

    def _is_valid_version_format(self, version: str) -> bool:
        """
        验证版本号格式是否有效

        Args:
            version (str): 版本号字符串

        Returns:
            bool: 版本号格式是否有效
        """
        try:
            from packaging import version as pkg_version
            pkg_version.parse(version)
            return True
        except Exception:
            return False

    @property
    def ui_manager(self):
        """
        获取UI管理器实例（延迟初始化）

        Returns:
            UpdateUIManager: UI管理器实例，如果不可用则返回None
        """
        if self._ui_manager is None and UI_AVAILABLE:
            try:
                import logging
                self._ui_manager = UpdateUIManager(self, self.parent)
                logging.getLogger(__name__).info("UI管理器初始化成功")
            except Exception as e:
                import logging
                logging.getLogger(__name__).error(f"UI管理器初始化失败: {e}")
                self._ui_manager = None

        return self._ui_manager

    def setup_update_ui(self, menu_bar, menu_title=None):
        """
        设置更新UI（标准接口）

        Args:
            menu_bar: 主窗口的菜单栏
            menu_title: 菜单标题，可选

        Returns:
            bool: 是否设置成功
        """
        import logging
        logger = logging.getLogger(__name__)

        try:
            if not UI_AVAILABLE:
                logger.warning("UI组件不可用，无法设置更新UI")
                return False

            ui_manager = self.ui_manager
            if ui_manager:
                ui_manager.setup_update_menu(menu_bar, menu_title)

                # 设置启动检查
                ui_manager.startup_update_check()

                logger.info("更新UI设置完成")
                return True
            else:
                logger.error("UI管理器不可用")
                return False

        except Exception as e:
            logger.error(f"设置更新UI失败: {e}")
            return False

    def force_check_updates_now(self) -> tuple:
        """
        立即强制检查更新（无视时间间隔限制）
        开发环境下的便利方法
        :return: (是否有更新, 下载路径, 本地版本, 消息)
        """
        print("强制检查更新中...")
        return self.check_for_updates(force_check=True)

    def check_for_updates_with_ui(self, force_check=True):
        """
        检查更新（带UI交互）

        Args:
            force_check: 是否强制检查更新
        """
        import logging
        logger = logging.getLogger(__name__)

        try:
            if self.ui_manager:
                self.ui_manager.check_for_updates_with_ui(force_check)
            else:
                logger.warning("UI管理器不可用，无法执行带UI的更新检查")

        except Exception as e:
            logger.error(f"带UI的更新检查失败: {e}")

    def show_about_dialog(self):
        """显示关于对话框"""
        import logging
        logger = logging.getLogger(__name__)

        try:
            if self.ui_manager:
                self.ui_manager.show_about_dialog()
            else:
                logger.warning("UI管理器不可用，无法显示关于对话框")

        except Exception as e:
            logger.error(f"显示关于对话框失败: {e}")

    def cleanup(self):
        """清理资源"""
        import logging
        logger = logging.getLogger(__name__)

        try:
            if self._ui_manager:
                self._ui_manager.cleanup()
                self._ui_manager = None

            logger.info("AutoUpdater 资源清理完成")

        except Exception as e:
            logger.error(f"清理AutoUpdater资源失败: {e}")

# 导入UI组件
try:
    from .ui import (
        UpdateUIManager,
        UpdateProgressDialog,
        AboutDialog,
        UpdateThread,
        UpdateStatusWidget,
        UpdateUIText,
        UpdateUIStyle,
        UpdateUIConfig
    )
    UI_AVAILABLE = True
except ImportError as e:
    import logging
    logger = logging.getLogger(__name__)
    logger.warning(f"更新UI组件导入失败: {e}")
    UI_AVAILABLE = False

# 导出的公共接口
__all__ = [
    # 核心更新功能
    'AutoUpdater',
    'GitHubClient',
    'DownloadManager',
    'BackupManager',
    'UpdateExecutor',
    'get_config',

    # UI组件（如果可用）
    'UpdateUIManager',
    'UpdateProgressDialog',
    'AboutDialog',
    'UpdateThread',
    'UpdateStatusWidget',

    # 资源和配置
    'UpdateUIText',
    'UpdateUIStyle',
    'UpdateUIConfig',
    'UI_AVAILABLE',

    # 异常类
    'UpdateError',
    'NetworkError',
    'VersionCheckError',
    'DownloadError',
    'BackupError',
    'UpdateExecutionError'
]