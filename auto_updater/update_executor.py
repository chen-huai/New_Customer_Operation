# -*- coding: utf-8 -*-
"""
更新执行模块
负责执行应用程序的热更新操作
"""

import os
import sys
import time
import subprocess
import shutil
import tempfile
from typing import Optional

from .config import (
    get_app_executable_path,
    get_executable_dir
)
from .config import get_config
from .backup_manager import BackupManager

# 异常类定义
class UpdateExecutionError(Exception):
    """更新执行异常"""
    pass

class UpdateExecutor:
    """更新执行器"""

    # 类变量：持久化存储延迟更新路径
    _delayed_update_path = None

    def __init__(self):
        self.config = get_config()
        self.backup_manager = BackupManager()

        # 导入两阶段更新器和自动完成器
        try:
            from .two_phase_updater import TwoPhaseUpdater
            self.two_phase_updater = TwoPhaseUpdater()
        except ImportError:
            self.two_phase_updater = None
            print("[更新器] 两阶段更新器不可用")

        try:
            from .auto_complete import AutoCompleter
            self.auto_completer = AutoCompleter()
        except ImportError:
            self.auto_completer = None
            print("[更新器] 自动完成器不可用")

    @property
    def delayed_update_path(self):
        """获取延迟更新路径"""
        return UpdateExecutor._delayed_update_path

    @delayed_update_path.setter
    def delayed_update_path(self, value):
        """设置延迟更新路径"""
        UpdateExecutor._delayed_update_path = value

    def execute_update(self, update_file_path: str, new_version: str) -> bool:
        """
        执行应用程序更新
        :param update_file_path: 更新文件路径
        :param new_version: 新版本号
        :return: 是否更新成功
        """
        try:
            # 验证更新文件
            if not os.path.exists(update_file_path):
                raise UpdateExecutionError("更新文件不存在")

            if os.path.getsize(update_file_path) == 0:
                raise UpdateExecutionError("更新文件无效")

            # 记录更新开始状态
            self._create_update_status_file("updating", new_version, update_file_path)

            # 获取当前可执行文件路径
            current_exe_path = get_app_executable_path()

            # 如果是开发环境，直接替换文件
            if not getattr(sys, 'frozen', False):
                return self._update_development_environment(update_file_path, new_version)

            # 生产环境（打包后的exe）需要特殊处理
            return self._update_production_environment(update_file_path, new_version)

        except UpdateExecutionError:
            # 记录更新失败状态
            self._create_update_status_file("failed", new_version, update_file_path)
            raise
        except Exception as e:
            # 记录更新失败状态
            self._create_update_status_file("failed", new_version, update_file_path)
            raise UpdateExecutionError(f"执行更新失败: {str(e)}")

    def _create_update_status_file(self, status: str, version: str, update_file: str) -> None:
        """
        创建更新状态文件

        Args:
            status: 更新状态 (updating/success/failed)
            version: 版本号
            update_file: 更新文件路径
        """
        try:
            import json
            from datetime import datetime

            status_data = {
                "status": status,
                "version": version,
                "update_file": update_file,
                "timestamp": datetime.now().isoformat(),
                "current_exe": get_app_executable_path()
            }

            status_file_path = os.path.join(get_executable_dir(), "update_status.json")
            with open(status_file_path, 'w', encoding='utf-8') as f:
                json.dump(status_data, f, indent=2, ensure_ascii=False)

            print(f"创建更新状态文件: {status}")

        except Exception as e:
            print(f"创建更新状态文件失败: {e}")

    def verify_update_success(self) -> bool:
        """
        验证更新是否成功

        Returns:
            更新是否成功
        """
        try:
            status_file_path = os.path.join(get_executable_dir(), "update_status.json")

            if not os.path.exists(status_file_path):
                print("未找到更新状态文件")
                return False

            import json
            with open(status_file_path, 'r', encoding='utf-8') as f:
                status_data = json.load(f)

            status = status_data.get("status")
            version = status_data.get("version")
            target_version = self.config.CURRENT_VERSION  # 使用配置常量

            print(f"更新状态: {status}, 版本: {version} -> {target_version}")

            # 检查更新状态
            if status != "success":
                print(f"更新状态不成功: {status}")
                return False

            # 检查版本是否匹配
            if version != target_version:
                print(f"版本不匹配: 期望 {version}, 实际 {target_version}")
                return False

            print("更新验证成功")
            return True

        except Exception as e:
            print(f"更新验证失败: {e}")
            return False

    def cleanup_update_status(self) -> None:
        """清理更新状态文件"""
        try:
            status_file_path = os.path.join(get_executable_dir(), "update_status.json")
            if os.path.exists(status_file_path):
                os.remove(status_file_path)
                print("清理更新状态文件")
        except Exception as e:
            print(f"清理更新状态文件失败: {e}")

    def _update_development_environment(self, update_file_path: str, new_version: str) -> bool:
        """
        开发环境下的更新（支持实际文件替换和版本更新）
        :param update_file_path: 更新文件路径
        :param new_version: 新版本号
        :return: 是否更新成功
        """
        try:
            print(f"开发环境更新开始: 版本 {new_version}")

            # 获取当前应用程序路径
            current_app_path = get_app_executable_path()
            print(f"当前应用程序路径: {current_app_path}")

            # 如果是Python文件，尝试创建备份
            if current_app_path.endswith('.py'):
                print("开发环境：检测到Python应用程序，创建备份...")
                backup_path = self.backup_manager.create_backup()
                if backup_path:
                    print(f"已创建开发环境备份: {os.path.basename(backup_path)}")

            # 强制执行文件替换（处理Python到exe的转换）
            try:
                print(f"开发环境：执行文件替换...")
                print(f"  更新文件: {update_file_path}")
                print(f"  当前文件: {current_app_path}")

                # 确定目标文件路径
                exec_dir = get_executable_dir()
                if update_file_path.endswith('.exe'):
                    # 如果更新文件是exe，目标应该是主程序目录下的exe
                    target_path = os.path.join(exec_dir, os.path.basename(update_file_path))
                else:
                    # 如果更新文件是py，目标是当前的py文件
                    target_path = current_app_path

                print(f"  目标文件: {target_path}")

                # 创建备份
                if os.path.exists(target_path):
                    backup_current = target_path + f".backup.{int(time.time())}"
                    shutil.copy2(target_path, backup_current)
                    print(f"  已创建备份: {os.path.basename(backup_current)}")

                # 强制替换文件
                shutil.copy2(update_file_path, target_path)
                print("  文件替换成功")

                # 验证替换后的文件
                if os.path.exists(target_path) and os.path.getsize(target_path) > 0:
                    print("  文件验证通过")
                    print(f"  新文件大小: {os.path.getsize(target_path)} bytes")
                else:
                    raise UpdateExecutionError("替换后的文件验证失败")

                # 如果成功替换为exe文件，记录这个信息供重启使用
                if update_file_path.endswith('.exe'):
                    print(f"  开发环境已更新为exe文件: {target_path}")
                    # 验证exe文件的可执行性
                    try:
                        if os.access(target_path, os.X_OK):
                            print("  exe文件可执行验证通过")
                        else:
                            print("  警告：exe文件可能不可执行")
                    except:
                        print("  无法验证exe文件可执行性")

                # 进行完整性验证
                if self._verify_file_replacement(update_file_path, target_path):
                    print("  文件替换完整性验证通过")
                else:
                    raise UpdateExecutionError("文件替换完整性验证失败")

            except Exception as file_error:
                print(f"文件替换失败: {file_error}")
                raise UpdateExecutionError(f"文件替换失败: {str(file_error)}")

            # 更新版本信息
            print("更新版本配置...")
            success = self.config.update_current_version(new_version)
            if success:
                # 记录更新成功状态
                self._create_update_status_file("success", new_version, update_file_path)
                print(f"版本信息已更新到: {new_version}")
                return True
            else:
                raise UpdateExecutionError("更新版本信息失败")

        except Exception as e:
            raise UpdateExecutionError(f"开发环境更新失败: {str(e)}")

    def _update_production_environment(self, update_file_path: str, new_version: str) -> bool:
        """
        生产环境下的更新（需要处理文件占用）
        :param update_file_path: 更新文件路径
        :param new_version: 新版本号
        :return: 是否更新成功
        """
        try:
            current_exe_path = get_app_executable_path()

            # 创建备份
            backup_path = self.backup_manager.create_backup()
            if not backup_path:
                raise UpdateExecutionError("创建备份失败")

            print(f"已创建备份: {backup_path}")

            # 尝试替换可执行文件
            replacement_success = self._replace_executable(update_file_path, current_exe_path)

            if replacement_success:
                # 验证文件替换的完整性
                if self._verify_file_replacement(update_file_path, current_exe_path):
                    # 更新版本文件
                    self.config.update_current_version(new_version)

                    # 记录更新成功状态
                    self._create_update_status_file("success", new_version, update_file_path)

                    print("文件替换成功，更新完成")
                    return True
                else:
                    print("文件替换验证失败，使用延迟更新")
                    # 回滚备份并使用延迟更新
                    self.backup_manager.restore_from_backup()
                    return self._schedule_delayed_update(update_file_path, current_exe_path, new_version)
            else:
                # 如果直接替换失败，使用批处理脚本延迟更新
                return self._schedule_delayed_update(update_file_path, current_exe_path, new_version)

        except Exception as e:
            raise UpdateExecutionError(f"生产环境更新失败: {str(e)}")

    def _replace_executable(self, source_path: str, target_path: str) -> bool:
        """
        替换可执行文件
        :param source_path: 源文件路径
        :param target_path: 目标文件路径
        :return: 是否替换成功
        """
        try:
            # 等待文件释放
            for _ in range(10):  # 最多等待10秒
                try:
                    # 尝试删除目标文件
                    if os.path.exists(target_path):
                        os.remove(target_path)

                    # 复制新文件
                    shutil.copy2(source_path, target_path)

                    # 验证文件是否正确复制
                    if os.path.exists(target_path) and os.path.getsize(target_path) > 0:
                        return True

                except PermissionError:
                    print("文件被占用，等待释放...")
                    time.sleep(1)
                except Exception as e:
                    print(f"替换文件失败: {e}")
                    time.sleep(1)

            return False

        except Exception as e:
            print(f"替换可执行文件失败: {e}")
            return False

    def _verify_file_replacement(self, source_path: str, target_path: str) -> bool:
        """
        验证文件替换的完整性

        Args:
            source_path: 源文件路径（更新文件）
            target_path: 目标文件路径（程序文件）

        Returns:
            文件替换是否成功
        """
        try:
            import hashlib

            # 检查文件是否存在
            if not os.path.exists(target_path):
                print("目标文件不存在")
                return False

            if not os.path.exists(source_path):
                print("源文件不存在")
                return False

            # 检查文件大小
            source_size = os.path.getsize(source_path)
            target_size = os.path.getsize(target_path)

            if source_size != target_size:
                print(f"文件大小不匹配: 源文件={source_size}, 目标文件={target_size}")
                return False

            if target_size == 0:
                print("文件大小为0，替换失败")
                return False

            # 检查文件哈希（对于小文件，否则只检查大小）
            if source_size < 50 * 1024 * 1024:  # 50MB以下的文件检查哈希
                try:
                    source_hash = self._calculate_file_hash(source_path)
                    target_hash = self._calculate_file_hash(target_path)

                    if source_hash != target_hash:
                        print("文件哈希不匹配")
                        return False

                    print("文件哈希验证通过")
                except Exception as e:
                    print(f"哈希验证失败，但大小匹配: {e}")

            print(f"文件替换验证通过: {target_size} bytes")
            return True

        except Exception as e:
            print(f"文件替换验证失败: {e}")
            return False

    def _force_copy_executable(self, source_path: str, target_dir: str) -> str:
        """
        强制复制可执行文件到目标目录

        Args:
            source_path: 源文件路径
            target_dir: 目标目录

        Returns:
            复制后的文件路径
        """
        try:
            if not os.path.exists(source_path):
                raise UpdateExecutionError(f"源文件不存在: {source_path}")

            # 确保目标目录存在
            os.makedirs(target_dir, exist_ok=True)

            # 获取源文件名
            source_filename = os.path.basename(source_path)
            target_path = os.path.join(target_dir, source_filename)

            # 如果目标文件已存在，删除它（强制覆盖）
            if os.path.exists(target_path):
                try:
                    os.remove(target_path)
                    print(f"删除旧文件: {target_path}")
                except PermissionError:
                    print(f"无法删除旧文件，可能被占用: {target_path}")

            # 复制新文件
            shutil.copy2(source_path, target_path)

            # 验证复制结果
            if os.path.exists(target_path) and os.path.getsize(target_path) > 0:
                print(f"文件复制成功: {target_path} ({os.path.getsize(target_path)} bytes)")
                return target_path
            else:
                raise UpdateExecutionError(f"文件复制验证失败: {target_path}")

        except Exception as e:
            raise UpdateExecutionError(f"强制复制文件失败: {str(e)}")

    def _calculate_file_hash(self, file_path: str) -> str:
        """
        计算文件的MD5哈希值

        Args:
            file_path: 文件路径

        Returns:
            MD5哈希值
        """
        hash_md5 = hashlib.md5()
        try:
            with open(file_path, "rb") as f:
                # 分块读取以处理大文件
                for chunk in iter(lambda: f.read(4096), b""):
                    hash_md5.update(chunk)
            return hash_md5.hexdigest()
        except Exception:
            raise UpdateExecutionError(f"计算文件哈希失败: {file_path}")

    def _schedule_delayed_update(self, update_file_path: str, current_exe_path: str, new_version: str) -> bool:
        """
        安排延迟更新（改进的两阶段更新方案）

        工作原理：
        阶段1: 创建待更新标记 → 保存新版本路径到文件 → 旧程序退出 → UI启动新版本
        阶段2: 新版本启动后 → 从文件读取路径 → 启动新版本 → 自动完成文件替换

        :param update_file_path: 更新文件路径（下载目录中的新版本）
        :param current_exe_path: 当前可执行文件路径
        :param new_version: 新版本号
        :return: 是否成功安排延迟更新
        """
        try:
            print(f"[两阶段更新] ========== 开始阶段1 ==========")
            print(f"[两阶段更新] 新版本路径: {update_file_path}")
            print(f"[两阶段更新] 目标路径: {current_exe_path}")
            print(f"[两阶段更新] 版本号: {new_version}")

            # ✅ 保存到类变量（供当前进程使用）
            self.delayed_update_path = update_file_path

            # ✅ 创建待更新标记（持久化到文件，重启后可读取）
            if self.two_phase_updater:
                success = self.two_phase_updater.create_pending_update(
                    update_file_path,
                    new_version
                )

                if success:
                    print(f"[两阶段更新] ✓ 待更新标记已创建")
                    print(f"[两阶段更新] ✓ 标记文件: {self.two_phase_updater.pending_marker_path}")
                    print(f"[两阶段更新] ✓ 新版本路径已保存到标记文件")
                    print(f"[两阶段更新] 阶段1完成，等待程序重启...")
                    print(f"[两阶段更新] ")
                    print(f"[两阶段更新] 更新流程:")
                    print(f"[两阶段更新] 1. UI将启动新版本: {update_file_path}")
                    print(f"[两阶段更新] 2. 新版本启动后，从标记文件读取路径")
                    print(f"[两阶段更新] 3. 启动新版本（从 downloads 目录）")
                    print(f"[两阶段更新] 4. 新版本自动检测并完成文件替换")
                    print(f"[两阶段更新] 5. 替换完成后，下次启动使用主目录exe")
                    print(f"[两阶段更新] ========== 阶段1完成 ==========")
                else:
                    raise UpdateExecutionError("创建待更新标记失败")
            else:
                # 降级方案：如果两阶段更新器不可用，使用原来的批处理脚本方案
                print(f"[两阶段更新] ⚠ 两阶段更新器不可用，使用批处理脚本方案")
                return self._schedule_delayed_update_fallback(update_file_path, current_exe_path, new_version)

            return True

        except Exception as e:
            error_msg = f"安排延迟更新失败: {str(e)}"
            print(f"[两阶段更新] ✗ {error_msg}")
            raise UpdateExecutionError(error_msg)

    def _schedule_delayed_update_fallback(self, update_file_path: str, current_exe_path: str, new_version: str) -> bool:
        """
        降级方案：使用批处理脚本（如果两阶段更新器不可用）

        这个方法保留原来的批处理脚本逻辑作为备用方案
        """
        try:
            import time
            from .config_constants import APP_NAME

            print(f"[批处理更新] 创建批处理脚本（降级方案）")

            # 生成唯一的脚本文件名
            timestamp = int(time.time())
            script_name = f"{APP_NAME}_update_{timestamp}.bat"
            script_path = os.path.join(tempfile.gettempdir(), script_name)

            # 创建简化的批处理脚本
            script_content = f'''@echo off
chcp 65001 >nul
echo 后台更新脚本启动
echo 等待程序退出...
timeout /t 5 /nobreak >nul
echo 开始文件替换...
copy /Y "{update_file_path}" "{current_exe_path}"
if %ERRORLEVEL% EQU 0 (
    echo 文件替换成功
) else (
    echo 文件替换失败
)
del "%~f0" 2>nul
'''

            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(script_content)

            # 启动更新脚本
            subprocess.Popen(
                [script_path],
                creationflags=subprocess.DETACHED_PROCESS,
                env=os.environ.copy(),
                encoding='utf-8'
            )

            print(f"[批处理更新] ✓ 脚本已启动: {script_path}")
            return True

        except Exception as e:
            raise UpdateExecutionError(f"创建批处理脚本失败: {str(e)}")

    def restart_application(self) -> bool:
        """
        重启应用程序
        :return: 是否成功重启
        """
        try:
            if getattr(sys, 'frozen', False):
                # 打包后的exe
                current_exe = sys.executable
                subprocess.Popen([current_exe],
                               env=os.environ.copy(),
                               encoding='utf-8')
            else:
                # 开发环境
                current_script = sys.argv[0]
                subprocess.Popen([sys.executable, current_script],
                               env=os.environ.copy(),
                               encoding='utf-8')

            # 退出当前进程
            sys.exit(0)

        except Exception as e:
            print(f"重启应用程序失败: {e}")
            return False

    def rollback_update(self) -> bool:
        """
        回滚到上一个版本
        :return: 是否回滚成功
        """
        try:
            # 获取最新备份
            latest_backup = self.backup_manager.get_latest_backup()
            if not latest_backup:
                raise UpdateExecutionError("没有找到可用的备份文件")

            # 从备份恢复
            success = self.backup_manager.restore_from_backup(latest_backup)
            if not success:
                raise UpdateExecutionError("从备份恢复失败")

            # 更新版本文件
            # 注意：这里需要根据实际情况确定如何获取备份的版本号
            # 暂时使用本地版本管理器的当前版本

            print("回滚成功")
            return True

        except Exception as e:
            raise UpdateExecutionError(f"回滚失败: {str(e)}")

    def validate_update_file(self, update_file_path: str) -> tuple:
        """
        验证更新文件的有效性
        :param update_file_path: 更新文件路径
        :return: (是否有效, 错误信息)
        """
        try:
            # 检查文件是否存在
            if not os.path.exists(update_file_path):
                return False, "更新文件不存在"

            # 检查文件大小
            file_size = os.path.getsize(update_file_path)
            if file_size == 0:
                return False, "更新文件为空"

            # 检查文件扩展名
            if not (update_file_path.endswith('.exe') or update_file_path.endswith('.zip')):
                return False, "更新文件格式不正确"

            # 基本可执行文件检查
            if update_file_path.endswith('.exe'):
                try:
                    # 检查PE文件头（简单检查）
                    with open(update_file_path, 'rb') as f:
                        header = f.read(2)
                        if header != b'MZ':  # DOS header
                            return False, "更新文件不是有效的可执行文件"
                except Exception as e:
                    return False, f"读取更新文件失败: {str(e)}"

            return True, "更新文件有效"

        except Exception as e:
            return False, f"验证更新文件失败: {str(e)}"

    def get_update_progress_info(self) -> dict:
        """
        获取更新进度信息
        :return: 进度信息字典
        """
        return {
            'is_updating': False,  # 是否正在更新
            'current_step': '',    # 当前步骤
            'progress': 0,         # 进度百分比
            'error_message': ''    # 错误信息
        }