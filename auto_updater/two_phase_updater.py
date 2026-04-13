# -*- coding: utf-8 -*-
"""
两阶段更新器 - 企业级解决方案

阶段1: 首次更新时运行新版本，记录待完成的文件替换
阶段2: 下次启动时完成文件替换，然后正常启动

优点：
- 无需替换正在运行的文件
- 平滑升级，用户无感知
- 支持回滚
- 可靠性高
"""
import os
import json
import shutil
import logging
from typing import Optional, Tuple
from datetime import datetime

from .config import get_executable_dir, get_app_executable_path
from .config_constants import APP_EXECUTABLE

logger = logging.getLogger(__name__)


class TwoPhaseUpdater:
    """
    两阶段更新器

    工作原理：
    阶段1: 下载新版本 → 运行新版本 → 创建待替换标记
    阶段2: 启动检查 → 完成文件替换 → 删除标记 → 正常运行
    """

    # 标记文件名
    PENDING_UPDATE_MARKER = ".pending_update.json"
    UPDATE_SUCCESS_MARKER = ".update_success.json"

    def __init__(self):
        self.exec_dir = get_executable_dir()
        self.pending_marker_path = os.path.join(self.exec_dir, self.PENDING_UPDATE_MARKER)
        self.success_marker_path = os.path.join(self.exec_dir, self.UPDATE_SUCCESS_MARKER)

    def create_pending_update(self, new_version_path: str, new_version: str) -> bool:
        """
        创建待更新标记（阶段1完成）

        Args:
            new_version_path: 新版本文件路径（下载目录中）
            new_version: 新版本号

        Returns:
            是否成功创建标记
        """
        try:
            target_path = get_app_executable_path()

            pending_data = {
                "source_file": new_version_path,
                "target_file": target_path,
                "version": new_version,
                "timestamp": datetime.now().isoformat(),
                "status": "pending"
            }

            # 写入标记文件
            with open(self.pending_marker_path, 'w', encoding='utf-8') as f:
                json.dump(pending_data, f, indent=2, ensure_ascii=False)

            logger.info(f"[两阶段更新] 创建待更新标记: {self.pending_marker_path}")
            logger.info(f"[两阶段更新] 源文件: {new_version_path}")
            logger.info(f"[两阶段更新] 目标文件: {target_path}")
            logger.info(f"[两阶段更新] 版本: {new_version}")

            return True

        except Exception as e:
            logger.error(f"[两阶段更新] 创建待更新标记失败: {e}")
            return False

    def has_pending_update(self) -> bool:
        """
        检查是否有待完成的更新

        Returns:
            是否存在待更新标记
        """
        return os.path.exists(self.pending_marker_path)

    def get_pending_update_info(self) -> Optional[dict]:
        """
        获取待更新信息

        Returns:
            待更新数据，如果不存在则返回None
        """
        try:
            if not os.path.exists(self.pending_marker_path):
                return None

            with open(self.pending_marker_path, 'r', encoding='utf-8') as f:
                return json.load(f)

        except Exception as e:
            logger.error(f"[两阶段更新] 读取待更新标记失败: {e}")
            return None

    def complete_pending_update(self) -> Tuple[bool, str]:
        """
        完成待更新（阶段2）

        在程序启动时调用，如果有待更新标记，则完成文件替换

        Returns:
            (是否成功, 消息)
        """
        try:
            # 读取待更新信息
            pending_data = self.get_pending_update_info()
            if not pending_data:
                return False, "没有待更新的数据"

            source_file = pending_data["source_file"]
            target_file = pending_data["target_file"]
            version = pending_data["version"]

            logger.info("=" * 60)
            logger.info("[两阶段更新] 阶段2: 完成文件替换")
            logger.info("=" * 60)
            logger.info(f"[两阶段更新] 源文件: {source_file}")
            logger.info(f"[两阶段更新] 目标文件: {target_file}")
            logger.info(f"[两阶段更新] 版本: {version}")

            # 验证源文件
            if not os.path.exists(source_file):
                error_msg = f"源文件不存在: {source_file}"
                logger.error(f"[两阶段更新] ✗ {error_msg}")
                # 删除无效的标记文件
                self.cleanup_pending_marker()
                return False, error_msg

            # 检查目标文件是否被占用
            if self._is_file_in_use(target_file):
                logger.warning(f"[两阶段更新] 目标文件仍在被占用")
                logger.info(f"[两阶段更新] 保持待更新状态，下次启动时重试")
                return False, "目标文件仍在被占用，稍后重试"

            # 备份目标文件
            backup_path = self._create_backup(target_file)
            if backup_path:
                logger.info(f"[两阶段更新] ✓ 已创建备份: {os.path.basename(backup_path)}")

            # 执行文件替换
            try:
                logger.info(f"[两阶段更新] 正在替换文件...")
                shutil.copy2(source_file, target_file)

                # 验证替换结果
                if os.path.exists(target_file):
                    file_size = os.path.getsize(target_file)
                    logger.info(f"[两阶段更新] ✓ 文件替换成功")
                    logger.info(f"[两阶段更新] 新文件大小: {file_size} bytes")

                    # 创建成功标记
                    self._create_success_marker(pending_data)

                    # 清理待更新标记
                    self.cleanup_pending_marker()

                    logger.info(f"[两阶段更新] ✓ 阶段2完成")
                    logger.info("=" * 60)

                    return True, f"更新成功: 版本 {version}"

                else:
                    raise Exception("替换后文件不存在")

            except Exception as e:
                logger.error(f"[两阶段更新] ✗ 文件替换失败: {e}")

                # 尝试恢复备份
                if backup_path and os.path.exists(backup_path):
                    logger.info(f"[两阶段更新] 尝试恢复备份...")
                    try:
                        shutil.copy2(backup_path, target_file)
                        logger.info(f"[两阶段更新] ✓ 备份已恢复")
                    except Exception as restore_error:
                        logger.error(f"[两阶段更新] ✗ 备份恢复失败: {restore_error}")

                return False, f"文件替换失败: {str(e)}"

        except Exception as e:
            error_msg = f"完成更新失败: {str(e)}"
            logger.error(f"[两阶段更新] ✗ {error_msg}")
            return False, error_msg

    def _is_file_in_use(self, file_path: str) -> bool:
        """
        检查文件是否正在被使用

        Args:
            file_path: 文件路径

        Returns:
            文件是否被占用
        """
        try:
            # 尝试以独占模式打开文件
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
            return False

        except Exception:
            # 文件被占用或其他错误
            return True

    def _create_backup(self, file_path: str) -> Optional[str]:
        """
        创建备份文件

        Args:
            file_path: 要备份的文件路径

        Returns:
            备份文件路径，失败返回None
        """
        try:
            from .backup_manager import BackupManager
            backup_manager = BackupManager()
            return backup_manager.create_backup()
        except Exception as e:
            logger.error(f"[两阶段更新] 创建备份失败: {e}")
            return None

    def _create_success_marker(self, update_data: dict) -> None:
        """
        创建更新成功标记

        Args:
            update_data: 更新数据
        """
        try:
            success_data = update_data.copy()
            success_data["status"] = "success"
            success_data["completed_at"] = datetime.now().isoformat()

            with open(self.success_marker_path, 'w', encoding='utf-8') as f:
                json.dump(success_data, f, indent=2, ensure_ascii=False)

            logger.info(f"[两阶段更新] 创建成功标记: {self.success_marker_path}")

        except Exception as e:
            logger.warning(f"[两阶段更新] 创建成功标记失败: {e}")

    def cleanup_pending_marker(self) -> bool:
        """
        清理待更新标记

        Returns:
            是否成功清理
        """
        try:
            if os.path.exists(self.pending_marker_path):
                os.remove(self.pending_marker_path)
                logger.info(f"[两阶段更新] 已清理待更新标记")
            return True
        except Exception as e:
            logger.error(f"[两阶段更新] 清理待更新标记失败: {e}")
            return False

    def cleanup_download_file(self, download_file_path: str) -> bool:
        """
        清理下载文件（在确认更新成功后调用）

        Args:
            download_file_path: 下载文件路径

        Returns:
            是否成功清理
        """
        try:
            if os.path.exists(download_file_path):
                os.remove(download_file_path)
                logger.info(f"[两阶段更新] 已清理下载文件: {download_file_path}")
            return True
        except Exception as e:
            logger.warning(f"[两阶段更新] 清理下载文件失败: {e}")
            return False


def check_and_complete_update_on_startup() -> Tuple[bool, str]:
    """
    程序启动时检查并完成更新（应该在main函数开始时调用）

    Returns:
        (是否完成更新, 消息)
    """
    updater = TwoPhaseUpdater()

    if not updater.has_pending_update():
        return False, "无待更新"

    logger.info("=" * 60)
    logger.info("[两阶段更新] 检测到待更新标记")
    logger.info("[两阶段更新] 开始执行阶段2...")
    logger.info("=" * 60)

    success, message = updater.complete_pending_update()

    if success:
        logger.info("[两阶段更新] ✓ 更新完成")
        # 可以在这里提示用户更新完成
    else:
        logger.warning(f"[两阶段更新] ⚠ {message}")

    return success, message
