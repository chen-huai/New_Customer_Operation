# -*- coding: utf-8 -*-
"""
备份管理模块
负责创建和恢复应用程序备份
"""

import os
import shutil
import time
import zipfile
from typing import Optional, List
from datetime import datetime

from .config import (
    get_backup_dir,
    get_app_executable_path,
    MAX_BACKUP_COUNT
)

# 异常类定义
class BackupError(Exception):
    """备份操作异常"""
    pass

class BackupManager:
    """备份管理器"""

    def __init__(self):
        self.backup_dir = get_backup_dir()
        os.makedirs(self.backup_dir, exist_ok=True)

    def create_backup(self, backup_name: Optional[str] = None) -> Optional[str]:
        """
        创建应用程序备份
        自动适应开发和生产环境
        :param backup_name: 备份名称（可选）
        :return: 备份文件路径，失败返回None
        """
        try:
            # 获取源文件路径
            app_path = get_app_executable_path()
            if not os.path.exists(app_path):
                raise BackupError(f"源文件不存在: {app_path}")

            # 确定环境类型和备份文件扩展名
            is_development = not app_path.endswith('.exe')
            backup_ext = '.py' if is_development else '.exe'

            # 生成备份名称
            if not backup_name:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                env_prefix = "dev" if is_development else "prod"
                backup_name = f"backup_{env_prefix}_{timestamp}"

            # 创建备份文件路径
            backup_file = os.path.join(self.backup_dir, f"{backup_name}{backup_ext}")

            # 如果备份文件已存在，删除它
            if os.path.exists(backup_file):
                os.remove(backup_file)

            # 复制文件
            shutil.copy2(app_path, backup_file)

            # 验证备份文件
            if not os.path.exists(backup_file) or os.path.getsize(backup_file) == 0:
                if os.path.exists(backup_file):
                    os.remove(backup_file)
                raise BackupError("备份文件创建失败")

            # 开发环境下额外备份相关配置文件
            if is_development:
                self._backup_development_files(backup_name)

            # 清理旧备份
            self._cleanup_old_backups()

            print(f"已创建{'开发环境' if is_development else '生产环境'}备份: {os.path.basename(backup_file)}")
            return backup_file

        except BackupError:
            raise
        except Exception as e:
            raise BackupError(f"创建备份失败: {str(e)}")

    def create_full_backup(self, backup_name: Optional[str] = None) -> Optional[str]:
        """
        创建完整备份（包含可执行文件和相关配置文件）
        :param backup_name: 备份名称（可选）
        :return: 备份文件路径，失败返回None
        """
        try:
            # 生成备份名称
            if not backup_name:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_name = f"full_backup_{timestamp}"

            backup_file = os.path.join(self.backup_dir, f"{backup_name}.zip")

            # 如果备份文件已存在，删除它
            if os.path.exists(backup_file):
                os.remove(backup_file)

            # 获取应用程序目录
            app_dir = os.path.dirname(get_app_executable_path())

            # 需要备份的文件列表
            files_to_backup = []

            # 主可执行文件
            exe_path = get_app_executable_path()
            if os.path.exists(exe_path):
                files_to_backup.append(exe_path)

            # 版本文件
            version_file = os.path.join(app_dir, "version.txt")
            if os.path.exists(version_file):
                files_to_backup.append(version_file)

            # 更新配置文件
            config_file = os.path.join(app_dir, "update_config.json")
            if os.path.exists(config_file):
                files_to_backup.append(config_file)

            if not files_to_backup:
                raise BackupError("没有找到需要备份的文件")

            # 创建ZIP备份文件
            with zipfile.ZipFile(backup_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file_path in files_to_backup:
                    if os.path.exists(file_path):
                        # 计算相对路径
                        arcname = os.path.basename(file_path)
                        zipf.write(file_path, arcname)

            # 验证备份文件
            if not os.path.exists(backup_file) or os.path.getsize(backup_file) == 0:
                if os.path.exists(backup_file):
                    os.remove(backup_file)
                raise BackupError("完整备份文件创建失败")

            # 清理旧备份
            self._cleanup_old_backups()

            return backup_file

        except BackupError:
            raise
        except Exception as e:
            raise BackupError(f"创建完整备份失败: {str(e)}")

    def restore_from_backup(self, backup_path: Optional[str] = None) -> bool:
        """
        从备份恢复应用程序
        :param backup_path: 备份文件路径（可选，默认使用最新备份）
        :return: 是否恢复成功
        """
        try:
            # 如果没有指定备份路径，使用最新的备份
            if not backup_path:
                backup_path = self.get_latest_backup()
                if not backup_path:
                    raise BackupError("没有找到可用的备份文件")

            # 检查备份文件是否存在
            if not os.path.exists(backup_path):
                raise BackupError(f"备份文件不存在: {backup_path}")

            # 获取目标路径
            app_path = get_app_executable_path()

            # 如果备份是ZIP文件，需要解压
            if backup_path.endswith('.zip'):
                return self._restore_from_zip_backup(backup_path, app_path)
            else:
                return self._restore_from_exe_backup(backup_path, app_path)

        except BackupError:
            raise
        except Exception as e:
            raise BackupError(f"恢复备份失败: {str(e)}")

    def _restore_from_exe_backup(self, backup_path: str, app_path: str) -> bool:
        """
        从EXE备份文件恢复
        :param backup_path: 备份文件路径
        :param app_path: 应用程序路径
        :return: 是否恢复成功
        """
        try:
            # 备份当前文件（如果存在）
            if os.path.exists(app_path):
                current_backup = os.path.join(
                    self.backup_dir,
                    f"current_backup_{int(time.time())}.exe"
                )
                shutil.copy2(app_path, current_backup)

            # 恢复文件
            shutil.copy2(backup_path, app_path)

            # 验证恢复的文件
            if not os.path.exists(app_path) or os.path.getsize(app_path) == 0:
                raise BackupError("恢复的文件无效")

            return True

        except Exception as e:
            raise BackupError(f"从EXE备份恢复失败: {str(e)}")

    def _restore_from_zip_backup(self, backup_path: str, app_path: str) -> bool:
        """
        从ZIP备份文件恢复
        :param backup_path: 备份文件路径
        :param app_path: 应用程序路径
        :return: 是否恢复成功
        """
        try:
            app_dir = os.path.dirname(app_path)

            # 解压备份文件
            with zipfile.ZipFile(backup_path, 'r') as zipf:
                zipf.extractall(app_dir)

            # 验证关键文件是否恢复成功
            if not os.path.exists(app_path):
                raise BackupError("主可执行文件恢复失败")

            return True

        except Exception as e:
            raise BackupError(f"从ZIP备份恢复失败: {str(e)}")

    def get_latest_backup(self) -> Optional[str]:
        """
        获取最新的备份文件路径
        :return: 最新备份文件路径，没有备份返回None
        """
        try:
            backups = self.list_backups()
            if backups:
                return backups[0]
            return None
        except Exception:
            return None

    def list_backups(self) -> List[str]:
        """
        列出所有备份文件，按修改时间排序（最新的在前）
        :return: 备份文件路径列表
        """
        try:
            backups = []
            for file_name in os.listdir(self.backup_dir):
                file_path = os.path.join(self.backup_dir, file_name)
                if os.path.isfile(file_path) and (
                    file_name.endswith('.exe') or file_name.endswith('.zip')
                ):
                    backups.append(file_path)

            # 按修改时间排序，最新的在前
            backups.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            return backups

        except Exception:
            return []

    def _backup_development_files(self, backup_name: str) -> None:
        """
        开发环境下额外备份相关文件
        :param backup_name: 备份名称基础
        """
        try:
            app_dir = os.path.dirname(get_app_executable_path())
            dev_files_to_backup = [
                'pdf_processor.py',
                'PDF_Rename_Operation.py',
                'PDF_Rename_UI.py',
                '打包工具.py',
                'PDF签名工具.py',
                'version_manager.py',
                'requirements.txt'
            ]

            for file_name in dev_files_to_backup:
                file_path = os.path.join(app_dir, file_name)
                if os.path.exists(file_path):
                    backup_file_path = os.path.join(self.backup_dir, f"{backup_name}_{file_name}")
                    if not os.path.exists(backup_file_path):
                        shutil.copy2(file_path, backup_file_path)
                        print(f"已备份开发文件: {file_name}")

        except Exception as e:
            print(f"备份开发文件时出现警告: {e}")

    def _cleanup_old_backups(self) -> bool:
        """
        清理旧备份文件，保留指定数量的最新备份
        :return: 是否清理成功
        """
        try:
            backups = self.list_backups()
            if len(backups) <= MAX_BACKUP_COUNT:
                return True

            # 删除多余的备份文件
            for backup_path in backups[MAX_BACKUP_COUNT:]:
                try:
                    os.remove(backup_path)
                except Exception as e:
                    print(f"删除备份文件失败 {backup_path}: {e}")

            return True

        except Exception as e:
            print(f"清理旧备份失败: {e}")
            return False

    def delete_backup(self, backup_path: str) -> bool:
        """
        删除指定的备份文件
        :param backup_path: 备份文件路径
        :return: 是否删除成功
        """
        try:
            if os.path.exists(backup_path):
                os.remove(backup_path)
                return True
            return False
        except Exception as e:
            print(f"删除备份失败: {e}")
            return False

    def get_backup_info(self, backup_path: str) -> dict:
        """
        获取备份文件信息
        :param backup_path: 备份文件路径
        :return: 备份信息字典
        """
        try:
            if not os.path.exists(backup_path):
                return {}

            stat = os.stat(backup_path)
            return {
                'path': backup_path,
                'name': os.path.basename(backup_path),
                'size': stat.st_size,
                'modified_time': datetime.fromtimestamp(stat.st_mtime),
                'created_time': datetime.fromtimestamp(stat.st_ctime),
                'is_zip': backup_path.endswith('.zip')
            }

        except Exception:
            return {}

    def verify_backup(self, backup_path: str) -> bool:
        """
        验证备份文件的完整性
        :param backup_path: 备份文件路径
        :return: 备份是否有效
        """
        try:
            if not os.path.exists(backup_path):
                return False

            if os.path.getsize(backup_path) == 0:
                return False

            # 如果是ZIP文件，检查是否可以正常读取
            if backup_path.endswith('.zip'):
                try:
                    with zipfile.ZipFile(backup_path, 'r') as zipf:
                        # 测试ZIP文件完整性
                        zipf.testzip()
                    return True
                except:
                    return False

            # EXE文件的基本检查
            return True

        except Exception:
            return False