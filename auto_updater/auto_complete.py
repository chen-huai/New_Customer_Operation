# -*- coding: utf-8 -*-
"""
自动完成更新的辅助工具

在新版本启动后，自动完成文件替换
"""
import os
import sys
import time
import json
import shutil
import logging
from threading import Thread, Event
from typing import Optional

from .config import get_executable_dir, get_app_executable_path
from .config_constants import APP_EXECUTABLE

logger = logging.getLogger(__name__)


class AutoCompleter:
    """
    自动完成更新器

    在新版本启动后，自动在后台完成文件替换
    """

    PENDING_UPDATE_MARKER = ".pending_update.json"

    def __init__(self):
        self.exec_dir = get_executable_dir()
        # 标记文件可能位于当前目录或父目录
        # 因为从 downloads 目录启动时，标记在主目录中
        self.marker_path = self._find_marker_file()
        self.is_running_from_download = self._check_if_running_from_download()

    def _find_marker_file(self) -> str:
        """
        查找标记文件

        标记文件可能在：
        1. 当前目录（如果从主目录启动）
        2. 父目录（如果从 downloads 目录启动）

        Returns:
            标记文件路径（如果存在）或默认路径
        """
        current_dir = self.exec_dir

        # 尝试当前目录
        marker_in_current = os.path.join(current_dir, self.PENDING_UPDATE_MARKER)
        if os.path.exists(marker_in_current):
            logger.info(f"[自动完成更新] 在当前目录找到标记文件: {marker_in_current}")
            return marker_in_current

        # 尝试父目录（处理从 downloads 子目录启动的情况）
        parent_dir = os.path.dirname(current_dir)
        if parent_dir and parent_dir != current_dir:  # 确保有父目录
            marker_in_parent = os.path.join(parent_dir, self.PENDING_UPDATE_MARKER)
            if os.path.exists(marker_in_parent):
                logger.info(f"[自动完成更新] 在父目录找到标记文件: {marker_in_parent}")
                return marker_in_parent

        # 都找不到，返回当前目录的路径（用于创建新标记）
        logger.info(f"[自动完成更新] 使用默认标记路径: {marker_in_current}")
        return marker_in_current

    def _check_if_running_from_download(self) -> bool:
        """
        检查当前是否从下载目录运行

        Returns:
            是否从下载目录运行
        """
        try:
            current_path = get_app_executable_path()

            # 方法1：直接检查路径中是否包含 "downloads" 目录
            # 这样无论 exec_dir 是什么，都能正确检测
            if "downloads" in current_path.lower():
                logger.info(f"[自动完成更新] 当前路径包含downloads: {current_path}")
                return True

            # 方法2：检查是否与待更新标记中的 source_file 一致
            pending_info = self.get_pending_update_info()
            if pending_info and "source_file" in pending_info:
                source_file = pending_info["source_file"]
                # 规范化路径后比较（处理大小写和路径分隔符）
                if os.path.normcase(os.path.normpath(current_path)) == os.path.normcase(os.path.normpath(source_file)):
                    logger.info(f"[自动完成更新] 当前运行文件与待更新源文件一致")
                    return True

            logger.info(f"[自动完成更新] 不是从下载目录运行: {current_path}")
            return False

        except Exception as e:
            logger.warning(f"[自动完成更新] 检测运行目录失败: {e}")
            return False

    def has_pending_update(self) -> bool:
        """检查是否有待完成的更新"""
        return os.path.exists(self.marker_path)

    def get_pending_update_info(self) -> Optional[dict]:
        """获取待更新信息"""
        try:
            if not os.path.exists(self.marker_path):
                return None
            with open(self.marker_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"读取待更新标记失败: {e}")
            return None

    def complete_update_in_background(self, callback=None):
        """
        在后台自动完成更新

        Args:
            callback: 完成回调函数 (success: bool, message: str) -> None
        """
        if not self.is_running_from_download:
            logger.info("[自动完成更新] 不是从下载目录运行，无需完成更新")
            return

        if not self.has_pending_update():
            logger.info("[自动完成更新] 没有待完成的更新")
            return

        # 创建后台线程完成更新
        thread = Thread(
            target=self._complete_update_thread,
            args=(callback,),
            daemon=True
        )
        thread.start()

        logger.info("[自动完成更新] 后台线程已启动")

    def _complete_update_thread(self, callback=None):
        """
        后台线程：完成文件替换

        Args:
            callback: 完成回调
        """
        try:
            logger.info("=" * 60)
            logger.info("[自动完成更新] 后台线程开始")
            logger.info("=" * 60)

            # 读取待更新信息
            pending_data = self.get_pending_update_info()
            if not pending_data:
                logger.error("[自动完成更新] 待更新标记不存在")
                if callback:
                    callback(False, "待更新标记不存在")
                return

            source_file = pending_data["source_file"]
            target_file = pending_data["target_file"]
            version = pending_data["version"]

            logger.info(f"[自动完成更新] 源文件: {source_file}")
            logger.info(f"[自动完成更新] 目标文件: {target_file}")
            logger.info(f"[自动完成更新] 版本: {version}")

            # 等待旧进程退出（最多等待30秒）
            logger.info(f"[自动完成更新] 等待旧进程退出...")
            for i in range(30):
                if not self._is_target_file_in_use(target_file):
                    logger.info(f"[自动完成更新] ✓ 旧进程已退出（等待{i}秒）")
                    break
                time.sleep(1)
            else:
                logger.warning(f"[自动完成更新] ⚠ 旧进程仍在运行，稍后重试")
                if callback:
                    callback(False, "旧进程仍在运行")
                return

            # 创建备份
            try:
                from .backup_manager import BackupManager
                backup_manager = BackupManager()
                backup_path = backup_manager.create_backup()
                if backup_path:
                    logger.info(f"[自动完成更新] ✓ 已创建备份")
            except Exception as e:
                logger.warning(f"[自动完成更新] 创建备份失败: {e}")
                backup_path = None

            # 执行文件替换
            try:
                logger.info(f"[自动完成更新] 正在替换文件...")
                shutil.copy2(source_file, target_file)

                # 验证替换结果
                if os.path.exists(target_file):
                    file_size = os.path.getsize(target_file)
                    logger.info(f"[自动完成更新] ✓ 文件替换成功")
                    logger.info(f"[自动完成更新] 新文件大小: {file_size} bytes")

                    # 清理标记
                    self._cleanup_marker()

                    # 清理下载文件（可选）
                    # self._cleanup_download_file(source_file)

                    logger.info("=" * 60)
                    logger.info("[自动完成更新] ✓ 更新完成")
                    logger.info("[自动完成更新] 下次启动将使用新版本")
                    logger.info("=" * 60)

                    if callback:
                        callback(True, f"更新完成: 版本 {version}")
                    return

                else:
                    raise Exception("替换后文件不存在")

            except Exception as e:
                logger.error(f"[自动完成更新] ✗ 文件替换失败: {e}")

                # 尝试恢复备份
                if backup_path and os.path.exists(backup_path):
                    logger.info(f"[自动完成更新] 尝试恢复备份...")
                    try:
                        shutil.copy2(backup_path, target_file)
                        logger.info(f"[自动完成更新] ✓ 备份已恢复")
                    except Exception as restore_error:
                        logger.error(f"[自动完成更新] ✗ 恢复失败: {restore_error}")

                if callback:
                    callback(False, f"文件替换失败: {str(e)}")
                return

        except Exception as e:
            logger.error(f"[自动完成更新] 后台线程异常: {e}")
            if callback:
                callback(False, f"后台更新异常: {str(e)}")

    def _is_target_file_in_use(self, file_path: str) -> bool:
        """
        检查目标文件是否被占用

        Args:
            file_path: 文件路径

        Returns:
            是否被占用
        """
        try:
            # 尝试打开文件
            if not os.path.exists(file_path):
                return False

            # Windows下检查文件占用
            if sys.platform == 'win32':
                try:
                    import win32file
                    import win32con

                    handle = win32file.CreateFile(
                        file_path,
                        win32con.GENERIC_READ,
                        0,  # 不共享
                        None,
                        win32con.OPEN_EXISTING,
                        0,
                        None
                    )
                    handle.Close()
                    return False  # 可以打开，说明未被占用
                except Exception:
                    return True  # 无法打开，说明被占用
            else:
                # Linux/Mac: 直接检查文件锁
                import fcntl
                with open(file_path, 'a') as f:
                    try:
                        fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)
                        fcntl.flock(f, fcntl.LOCK_UN)
                        return False
                    except IOError:
                        return True

        except Exception as e:
            logger.warning(f"检查文件占用失败: {e}")
            return True  # 出错时假设被占用，更安全

    def _cleanup_marker(self):
        """清理待更新标记"""
        try:
            if os.path.exists(self.marker_path):
                os.remove(self.marker_path)
                logger.info("[自动完成更新] ✓ 已清理待更新标记")
        except Exception as e:
            logger.error(f"[自动完成更新] 清理标记失败: {e}")

    def _cleanup_download_file(self, download_file: str):
        """
        清理下载文件（可选）

        Args:
            download_file: 下载文件路径
        """
        try:
            if os.path.exists(download_file):
                os.remove(download_file)
                logger.info(f"[自动完成更新] ✓ 已清理下载文件")
        except Exception as e:
            logger.warning(f"[自动完成更新] 清理下载文件失败: {e}")


def auto_complete_update_if_needed(callback=None):
    """
    如果需要，自动完成更新

    在新版本启动后调用此函数，会自动在后台完成文件替换

    Args:
        callback: 完成回调函数 (success: bool, message: str) -> None

    Returns:
        是否已启动后台更新
    """
    completer = AutoCompleter()

    if completer.is_running_from_download and completer.has_pending_update():
        logger.info("[自动完成更新] 检测到需要完成更新")
        completer.complete_update_in_background(callback)
        return True
    else:
        logger.info("[自动完成更新] 无需完成更新")
        return False
