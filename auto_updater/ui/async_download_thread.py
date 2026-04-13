"""
Async Download Thread

Provides asynchronous download functionality using QThread to prevent UI blocking.
Supports progress updates, cancellation, and error handling through signals.
"""

import time
import threading
from typing import Optional, Tuple
from PyQt5.QtCore import QThread, pyqtSignal
import requests
import logging

logger = logging.getLogger(__name__)


class AsyncDownloadThread(QThread):
    """异步下载线程类"""

    # 信号定义
    progress_updated = pyqtSignal(int, int, float)  # 已下载, 总大小, 百分比
    download_complete = pyqtSignal(bool, str, str)  # 是否成功, 文件路径, 错误信息
    error_occurred = pyqtSignal(str)  # 错误信息
    status_changed = pyqtSignal(str)  # 状态变更
    speed_calculated = pyqtSignal(float)  # 下载速度 MB/s
    retry_attempt = pyqtSignal(int, int)  # 重试次数, 最大重试次数

    def __init__(self, url: str, file_path: str, headers: dict = None, max_retries: int = 3):
        """
        初始化异步下载线程

        Args:
            url: 下载链接
            file_path: 保存路径
            headers: 请求头
            max_retries: 最大重试次数
        """
        super().__init__()
        self.url = url
        self.file_path = file_path
        self.headers = headers or {}
        self.max_retries = max_retries
        self.should_stop = False
        self.is_paused = False

        # 下载状态
        self.start_time = None
        self.last_progress_time = None
        self.last_downloaded = 0
        self.current_retry = 0

        logger.info(f"异步下载线程初始化: {url} -> {file_path}, 最大重试: {max_retries}")

    def run(self):
        """执行下载操作（带重试机制）"""
        try:
            self.should_stop = False
            self.is_paused = False
            self.start_time = time.time()
            self.last_progress_time = self.start_time
            self.last_downloaded = 0
            self.current_retry = 0

            logger.info(f"开始异步下载: {self.url}, 最大重试: {self.max_retries}")
            self.status_changed.emit("正在连接服务器...")

            # 执行带重试的下载
            for attempt in range(self.max_retries + 1):  # +1 因为第一次不算重试
                if self.should_stop:
                    logger.info("下载被用户取消")
                    self.download_complete.emit(False, "", "下载已被取消")
                    return

                if attempt > 0:
                    self.current_retry = attempt
                    self.retry_attempt.emit(attempt, self.max_retries)

                    # 智能重试延迟策略
                    delay = self._calculate_retry_delay(error, attempt)
                    self.status_changed.emit(f"网络重试 {attempt}/{self.max_retries}，{delay:.1f}秒后开始...")
                    logger.info(f"重试第 {attempt} 次，等待 {delay:.1f} 秒")

                    # 在等待期间检查取消（优化检查频率）
                    check_interval = min(0.1, delay / 20)  # 动态调整检查间隔
                    elapsed = 0.0
                    while elapsed < delay:
                        if self.should_stop:
                            self.download_complete.emit(False, "", "下载已被取消")
                            return
                        time.sleep(check_interval)
                        elapsed += check_interval

                success, file_path, error = self._download_file()

                if success or self.should_stop:
                    # 下载成功或被用户取消
                    if self.should_stop:
                        logger.info("下载被用户取消")
                        self.download_complete.emit(False, "", "下载已被取消")
                    else:
                        logger.info(f"下载完成: success={success}, file_path={file_path}")
                        self.download_complete.emit(success, file_path or "", error or "")
                    return
                else:
                    # 下载失败，准备重试
                    logger.warning(f"下载失败（尝试 {attempt + 1}/{self.max_retries + 1}）: {error}")
                    if attempt < self.max_retries:
                        continue  # 继续重试
                    else:
                        # 所有重试都失败了
                        final_error = f"下载失败，已重试{self.max_retries}次: {error}"
                        logger.error(final_error)
                        self.error_occurred.emit(final_error)
                        self.download_complete.emit(False, "", final_error)
                        return

        except Exception as e:
            logger.error(f"异步下载异常: {e}", exc_info=True)
            self.error_occurred.emit(f"下载异常: {str(e)}")
            self.download_complete.emit(False, "", f"下载异常: {str(e)}")

    def _calculate_retry_delay(self, error: str, attempt: int) -> float:
        """
        计算智能重试延迟

        Args:
            error: 上次的错误信息
            attempt: 当前重试次数

        Returns:
            float: 延迟秒数
        """
        error_lower = error.lower() if error else ""

        # 基于错误类型的不同策略
        if any(keyword in error_lower for keyword in ["timeout", "超时"]):
            # 超时错误：较短的初始延迟，快速重试
            base_delay = 1.0
            max_delay = 15.0
        elif any(keyword in error_lower for keyword in ["connection", "连接"]):
            # 连接错误：中等延迟
            base_delay = 2.0
            max_delay = 20.0
        elif any(keyword in error_lower for keyword in ["dns", "resolve", "解析"]):
            # DNS解析错误：较长延迟
            base_delay = 3.0
            max_delay = 30.0
        else:
            # 未知错误：默认策略
            base_delay = 2.0
            max_delay = 25.0

        # 指数退避，但有上限
        delay = min(base_delay * (1.5 ** attempt), max_delay)

        # 添加小随机抖动，避免同时重试
        import random
        jitter = random.uniform(0.8, 1.2)
        delay *= jitter

        return max(0.5, delay)  # 最小延迟0.5秒

    def _classify_error(self, error: str) -> str:
        """
        分类错误类型

        Args:
            error: 错误信息

        Returns:
            str: 错误类型 (timeout/connection/dns/other)
        """
        if not error:
            return "other"

        error_lower = error.lower()
        if any(keyword in error_lower for keyword in ["timeout", "超时"]):
            return "timeout"
        elif any(keyword in error_lower for keyword in ["connection", "连接"]):
            return "connection"
        elif any(keyword in error_lower for keyword in ["dns", "resolve", "解析"]):
            return "dns"
        else:
            return "other"

    def run(self):
        """执行下载操作（带重试机制）"""
        try:
            self.should_stop = False
            self.is_paused = False
            self.start_time = time.time()
            self.last_progress_time = self.start_time
            self.last_downloaded = 0
            self.current_retry = 0

            logger.info(f"开始异步下载: {self.url}, 最大重试: {self.max_retries}")
            self.status_changed.emit("正在连接服务器...")

            # 执行带重试的下载
            error = None
            for attempt in range(self.max_retries + 1):  # +1 因为第一次不算重试
                if self.should_stop:
                    logger.info("下载被用户取消")
                    self.download_complete.emit(False, "", "下载已被取消")
                    return

                if attempt > 0:
                    self.current_retry = attempt
                    self.retry_attempt.emit(attempt, self.max_retries)

                    # 智能重试延迟策略
                    delay = self._calculate_retry_delay(error, attempt)
                    self.status_changed.emit(f"网络重试 {attempt}/{self.max_retries}，{delay:.1f}秒后开始...")
                    logger.info(f"重试第 {attempt} 次，等待 {delay:.1f} 秒")

                    # 在等待期间检查取消（优化检查频率）
                    check_interval = min(0.1, delay / 20)  # 动态调整检查间隔
                    elapsed = 0.0
                    while elapsed < delay:
                        if self.should_stop:
                            self.download_complete.emit(False, "", "下载已被取消")
                            return
                        time.sleep(check_interval)
                        elapsed += check_interval

                success, file_path, error = self._download_file()

                if success or self.should_stop:
                    # 下载成功或被用户取消
                    if self.should_stop:
                        logger.info("下载被用户取消")
                        self.download_complete.emit(False, "", "下载已被取消")
                    else:
                        logger.info(f"下载完成: success={success}, file_path={file_path}")
                        self.download_complete.emit(success, file_path or "", error or "")
                    return
                else:
                    # 下载失败，准备重试
                    logger.warning(f"下载失败（尝试 {attempt + 1}/{self.max_retries + 1}）: {error}")
                    if attempt < self.max_retries:
                        continue  # 继续重试
                    else:
                        # 所有重试都失败了
                        final_error = f"下载失败，已重试{self.max_retries}次: {error}"
                        logger.error(final_error)
                        self.error_occurred.emit(final_error)
                        self.download_complete.emit(False, "", final_error)
                        return

        except Exception as e:
            logger.error(f"异步下载异常: {e}", exc_info=True)
            self.error_occurred.emit(f"下载异常: {str(e)}")
            self.download_complete.emit(False, "", f"下载异常: {str(e)}")

    def _download_file(self) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        执行实际的文件下载

        Returns:
            tuple: (是否成功, 文件路径, 错误信息)
        """
        try:
            # 创建会话
            session = requests.Session()
            session.headers.update(self.headers)

            # 导入超时配置
            from ..config_constants import NETWORK_TIMEOUT_SHORT, NETWORK_TIMEOUT_LONG

            # 发送HEAD请求获取文件信息
            try:
                response = session.head(self.url, timeout=NETWORK_TIMEOUT_SHORT)
                response.raise_for_status()
                total_size = int(response.headers.get('content-length', 0))
            except requests.RequestException as e:
                logger.warning(f"HEAD请求失败，改用GET请求: {e}")
                response = session.get(self.url, stream=True, timeout=NETWORK_TIMEOUT_SHORT)
                total_size = int(response.headers.get('content-length', 0))

            if total_size == 0:
                logger.warning("无法获取文件大小，继续下载")

            self.status_changed.emit("开始下载...")

            # 开始下载
            with session.get(self.url, stream=True, timeout=NETWORK_TIMEOUT_LONG) as response:
                response.raise_for_status()

                downloaded = 0
                with open(self.file_path, 'wb') as file:
                    for chunk in response.iter_content(chunk_size=8192):
                        if self.should_stop:
                            logger.info("检测到停止信号，终止下载")
                            return False, None, "下载已被取消"

                        if self.is_paused:
                            while self.is_paused and not self.should_stop:
                                time.sleep(0.1)
                            continue

                        if chunk:
                            file.write(chunk)
                            downloaded += len(chunk)

                            # 计算并发送进度更新
                            current_time = time.time()
                            if current_time - self.last_progress_time >= 0.1:  # 每100ms更新一次
                                percentage = (downloaded / total_size * 100) if total_size > 0 else 0
                                self.progress_updated.emit(downloaded, total_size, percentage)

                                # 计算下载速度
                                time_diff = current_time - self.last_progress_time
                                if time_diff > 0:
                                    speed_bytes = (downloaded - self.last_downloaded) / time_diff
                                    speed_mbs = speed_bytes / (1024 * 1024)
                                    self.speed_calculated.emit(speed_mbs)

                                self.last_progress_time = current_time
                                self.last_downloaded = downloaded

            # 下载完成
            final_percentage = 100.0 if total_size > 0 else 100.0
            self.progress_updated.emit(downloaded, total_size, final_percentage)

            logger.info(f"文件下载成功: {self.file_path}, 大小: {downloaded} bytes")
            return True, self.file_path, None

        except requests.exceptions.Timeout:
            error_msg = "下载超时，请检查网络连接"
            logger.error(error_msg)
            return False, None, error_msg
        except requests.exceptions.ConnectionError:
            error_msg = "网络连接失败，请检查网络设置"
            logger.error(error_msg)
            return False, None, error_msg
        except requests.exceptions.RequestException as e:
            error_msg = f"网络请求失败: {str(e)}"
            logger.error(error_msg)
            return False, None, error_msg
        except IOError as e:
            error_msg = f"文件写入失败: {str(e)}"
            logger.error(error_msg)
            return False, None, error_msg
        except Exception as e:
            error_msg = f"下载过程中发生未知错误: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return False, None, error_msg

    def cancel_download(self):
        """取消下载"""
        logger.info("请求取消下载")
        self.should_stop = True
        self.is_paused = False

    def pause_download(self):
        """暂停下载"""
        logger.info("暂停下载")
        self.is_paused = True

    def resume_download(self):
        """恢复下载"""
        logger.info("恢复下载")
        self.is_paused = False

    def is_downloading(self) -> bool:
        """检查是否正在下载"""
        return self.isRunning() and not self.should_stop

    def get_download_info(self) -> dict:
        """获取下载信息"""
        return {
            "url": self.url,
            "file_path": self.file_path,
            "is_running": self.isRunning(),
            "should_stop": self.should_stop,
            "is_paused": self.is_paused,
            "start_time": self.start_time
        }