# -*- coding: utf-8 -*-
"""
文件下载管理模块
负责下载更新文件、验证文件完整性和显示下载进度
"""

import os
import requests
import hashlib
import time
from typing import Optional, Callable
from urllib.parse import urlparse

from .config import (
    DOWNLOAD_TIMEOUT,
    get_executable_dir,
    APP_NAME,
    APP_EXECUTABLE,
    REQUEST_HEADERS,
    CONNECTION_TIMEOUT,
    RETRY_DELAY,
    MAX_RETRIES
)
from .retry_utils import RetryExecutor, DefaultRetryStrategy
from .config_constants import NetworkConfig
import logging

logger = logging.getLogger(__name__)

# 异常类定义
class DownloadError(Exception):
    """文件下载异常"""
    pass

class DownloadManager:
    """文件下载管理器"""

    def __init__(self):
        self.session = requests.Session()
        # 使用统一的请求头配置
        self.session.headers.update(REQUEST_HEADERS)

        # 配置连接池和重试策略
        self._configure_session()

        # 初始化重试执行器（下载使用较保守的策略）
        retry_strategy = DefaultRetryStrategy(max_retries=2, base_delay=RETRY_DELAY)
        self.retry_executor = RetryExecutor(retry_strategy)

    def _configure_session(self):
        """配置requests会话"""
        from requests.adapters import HTTPAdapter
        from urllib3.util.retry import Retry

        # 配置重试策略（下载使用更保守的重试策略）
        retry_strategy = Retry(
            total=NetworkConfig.CONNECTION_POOL['max_retries'],
            backoff_factor=NetworkConfig.CONNECTION_POOL['backoff_factor'],
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["HEAD", "GET", "OPTIONS"]
        )

        adapter = HTTPAdapter(
            max_retries=retry_strategy,
            pool_connections=5,  # 下载使用较少的连接池
            pool_maxsize=10
        )
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)

    def _calculate_file_hash(self, file_path: str) -> str:
        """
        计算文件的SHA256哈希值
        :param file_path: 文件路径
        :return: SHA256哈希值
        """
        sha256_hash = hashlib.sha256()
        try:
            with open(file_path, "rb") as f:
                # 分块读取文件以避免内存问题
                for chunk in iter(lambda: f.read(4096), b""):
                    sha256_hash.update(chunk)
            return sha256_hash.hexdigest()
        except Exception as e:
            raise DownloadError(f"计算文件哈希失败: {str(e)}")

    def _verify_file_integrity(self, file_path: str, expected_hash: Optional[str] = None) -> bool:
        """
        验证文件完整性
        :param file_path: 文件路径
        :param expected_hash: 期望的哈希值（可选）
        :return: 文件是否完整
        """
        try:
            if not os.path.exists(file_path):
                return False

            # 检查文件大小
            if os.path.getsize(file_path) == 0:
                return False

            # 如果提供了期望的哈希值，进行哈希验证
            if expected_hash:
                calculated_hash = self._calculate_file_hash(file_path)
                return calculated_hash.lower() == expected_hash.lower()

            # 基本完整性检查：文件不为空且可以读取
            try:
                with open(file_path, 'rb') as f:
                    f.read(1)  # 尝试读取第一个字节
                return True
            except Exception:
                return False

        except Exception as e:
            raise DownloadError(f"验证文件完整性失败: {str(e)}")

    def download_file(self, url: str, version: str, progress_callback: Optional[Callable] = None) -> Optional[str]:
        """
        下载文件
        :param url: 下载链接
        :param version: 版本号
        :param progress_callback: 进度回调函数 (downloaded, total, percentage)
        :return: 下载的文件路径，失败返回None
        """
        try:
            # 验证URL
            parsed_url = urlparse(url)
            if not parsed_url.scheme or not parsed_url.netloc:
                raise DownloadError("无效的下载链接")

            # 创建下载目录
            download_dir = os.path.join(get_executable_dir(), "downloads")
            os.makedirs(download_dir, exist_ok=True)

            # 生成下载文件名（不带版本号）
            file_name = APP_EXECUTABLE  # 直接使用可执行文件名
            file_path = os.path.join(download_dir, file_name)

            # 如果文件已存在，先删除
            if os.path.exists(file_path):
                os.remove(file_path)

            # 开始下载
            try:
                response = self.session.get(
                    url,
                    stream=True,
                    timeout=(CONNECTION_TIMEOUT, DOWNLOAD_TIMEOUT)  # 连接超时 + 读取超时
                )
                response.raise_for_status()

                # 获取文件总大小（改进的错误处理）
                try:
                    content_length = response.headers.get('content-length')
                    logger.debug(f"Content-Length响应头: {content_length}")

                    if content_length is not None and content_length.strip():
                        total_size = int(content_length)
                        logger.info(f"文件大小: {total_size} bytes ({total_size/1024/1024:.1f} MB)")

                        # 验证文件大小的合理性（1MB - 10GB范围）
                        if total_size < 0:
                            logger.warning(f"无效的文件大小: {total_size}，将使用动态计算")
                            total_size = 0
                        elif total_size > 10 * 1024 * 1024 * 1024:  # 10GB
                            logger.warning(f"文件大小过大: {total_size} bytes，可能不准确")
                        elif total_size < 1024 * 1024:  # 1MB
                            logger.info(f"文件大小较小: {total_size} bytes ({total_size/1024:.1f} KB)")
                        else:
                            logger.info(f"文件大小正常: {total_size} bytes ({total_size/1024/1024:.1f} MB)")
                    else:
                        logger.info("服务器未提供content-length，将使用动态进度显示")
                        total_size = 0
                except (ValueError, TypeError, AttributeError) as e:
                    logger.warning(f"解析content-length失败: {e}，原始值: '{content_length}'，将使用动态计算")
                    total_size = 0

                downloaded = 0

                # 写入文件
                with open(file_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                            chunk_size = len(chunk)
                            downloaded += chunk_size

                            # 调用进度回调（改进的进度计算）
                            if progress_callback:
                                try:
                                    if total_size > 0:
                                        # 已知总大小的计算
                                        percentage = min(100, (downloaded / total_size * 100))
                                        # 添加调试日志（每10MB记录一次）
                                        if downloaded % (10 * 1024 * 1024) < chunk_size:
                                            logger.debug(f"下载进度: {downloaded}/{total_size} bytes ({percentage:.1f}%)")
                                    else:
                                        # 未知总大小时的估算（基于已下载大小和经验值）
                                        # 假设文件大小约为200MB，用于显示相对进度
                                        estimated_size = 200 * 1024 * 1024  # 200MB
                                        percentage = min(100, (downloaded / estimated_size * 50))  # 最多显示50%

                                    # 确保百分比在合理范围内
                                    percentage = max(0, min(100, int(percentage)))
                                    progress_callback(downloaded, total_size, percentage)

                                except (ValueError, TypeError, ZeroDivisionError) as e:
                                    logger.warning(f"进度计算异常: {e}")
                                    # 发送安全的默认进度
                                    progress_callback(downloaded, total_size, 0)

                # 下载完成，记录最终状态
                actual_file_size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
                logger.info(f"下载完成 - 预期大小: {total_size}, 实际大小: {actual_file_size}, 下载位置: {file_path}")

                # 如果有原始大小但与实际大小不匹配，记录警告
                if total_size > 0 and actual_file_size != total_size:
                    logger.warning(f"文件大小不匹配 - 预期: {total_size}, 实际: {actual_file_size}")
                elif total_size == 0:
                    logger.info(f"动态下载完成 - 实际大小: {actual_file_size} bytes")

                # 验证下载的文件
                if not self._verify_file_integrity(file_path):
                    if os.path.exists(file_path):
                        os.remove(file_path)
                    raise DownloadError("下载的文件完整性验证失败")

                logger.info("文件完整性验证通过")
                return file_path

            except requests.exceptions.Timeout:
                raise DownloadError(f"下载超时（{DOWNLOAD_TIMEOUT}秒）")
            except requests.exceptions.ConnectionError as e:
                error_msg = str(e).lower()
                if "name resolution failed" in error_msg or "getaddrinfo failed" in error_msg:
                    raise DownloadError("DNS解析失败，请检查网络连接")
                elif "connection refused" in error_msg:
                    raise DownloadError("连接被拒绝，可能是网络问题")
                elif "ssl" in error_msg or "certificate" in error_msg:
                    raise DownloadError("SSL证书验证失败")
                else:
                    raise DownloadError("网络连接失败，请检查网络设置")
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 404:
                    raise DownloadError("下载文件不存在")
                elif e.response.status_code == 403:
                    raise DownloadError("访问被拒绝，可能需要权限")
                elif e.response.status_code == 429:
                    raise DownloadError("请求频率限制，请稍后重试")
                else:
                    raise DownloadError(f"服务器错误 {e.response.status_code}")
            except requests.exceptions.RequestException as e:
                raise DownloadError(f"下载请求异常: {str(e)}")

        except DownloadError:
            raise
        except Exception as e:
            raise DownloadError(f"下载失败: {str(e)}")

    def download_with_retry(self, url: str, version: str, max_retries: int = None,
                           progress_callback: Optional[Callable] = None) -> Optional[str]:
        """
        带重试机制的文件下载
        :param url: 下载链接
        :param version: 版本号
        :param max_retries: 最大重试次数（默认使用配置值）
        :param progress_callback: 进度回调函数
        :return: 下载的文件路径，失败返回None
        """
        if max_retries is None:
            max_retries = MAX_RETRIES

        last_error = None

        for attempt in range(max_retries):
            try:
                if attempt > 0:
                    # 更合理的重试延迟：2, 4, 8秒
                    wait_time = min(RETRY_DELAY * (2 ** (attempt - 1)), 15)

                    # 通知UI进入等待状态（使用特殊的进度值）
                    if progress_callback:
                        try:
                            # 使用-1作为等待状态的标识符，避免与实际进度混淆
                            progress_callback(0, 0, -1)
                            # 通过另一个回调通知等待时间（如果UI支持的话）
                            logger.info(f"重试前等待 {wait_time} 秒...")
                        except Exception as e:
                            logger.warning(f"进度回调通知失败: {e}")

                    print(f"等待 {wait_time} 秒后重试...")
                    time.sleep(wait_time)

                logger.info(f"下载尝试 {attempt + 1}/{max_retries}")
                file_path = self.download_file(url, version, progress_callback)
                if file_path:
                    logger.info(f"下载成功: {file_path}")
                    return file_path

            except DownloadError as e:
                last_error = e
                error_msg = str(e).lower()

                # 分析错误类型，决定是否应该重试
                should_retry = False
                if "超时" in error_msg:
                    should_retry = True
                elif "网络连接失败" in error_msg and "dns" not in error_msg and "ssl" not in error_msg:
                    should_retry = True
                elif "服务器错误 5" in error_msg:  # 5xx服务器错误
                    should_retry = True
                elif "请求频率限制" in error_msg:
                    should_retry = True
                    # 对于频率限制，延长等待时间
                    time.sleep(30)

                if attempt == max_retries - 1:
                    # 最后一次尝试，不再继续重试
                    break

                if not should_retry:
                    logger.warning(f"错误类型不适合重试: {e}")
                    break

                logger.warning(f"下载失败（可重试）: {e}")

        # 所有重试都失败了
        if last_error:
            raise DownloadError(f"下载失败，已重试{max_retries}次: {str(last_error)}")
        else:
            raise DownloadError("下载失败，原因未知")

    def cleanup_downloads(self, keep_count: int = 2) -> bool:
        """
        清理下载目录，保留最近的几个文件
        :param keep_count: 保留文件数量
        :return: 是否清理成功
        """
        try:
            download_dir = os.path.join(get_executable_dir(), "downloads")
            if not os.path.exists(download_dir):
                return True

            # 获取所有下载文件（清理所有相关版本的可执行文件）
            files = []
            base_name = os.path.splitext(APP_EXECUTABLE)[0]  # 去掉.exe扩展名
            current_name = APP_EXECUTABLE
            old_prefix_v = f"{base_name}_v"  # 带版本号的旧文件名
            old_prefix_dot = f"{APP_NAME}.v"  # 旧的文件名前缀

            for file_name in os.listdir(download_dir):
                # 清理当前可执行文件名、带版本号的文件和旧格式文件
                if (file_name == current_name or
                    file_name.startswith(old_prefix_v) or
                    file_name.startswith(old_prefix_dot)):
                    file_path = os.path.join(download_dir, file_name)
                    if os.path.isfile(file_path):
                        files.append((file_path, os.path.getmtime(file_path)))

            # 按修改时间排序，保留最新的文件
            files.sort(key=lambda x: x[1], reverse=True)

            # 删除旧文件
            for file_path, _ in files[keep_count:]:
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"删除文件失败 {file_path}: {e}")

            return True

        except Exception as e:
            print(f"清理下载目录失败: {e}")
            return False

    def get_download_size(self, url: str) -> Optional[int]:
        """
        获取下载文件大小
        :param url: 下载链接
        :return: 文件大小（字节），获取失败返回None
        """
        try:
            response = self.session.head(url, timeout=30)
            response.raise_for_status()
            return int(response.headers.get('content-length', 0))
        except Exception:
            return None

    def test_download_speed(self, url: str) -> float:
        """
        测试下载速度
        :param url: 测试URL
        :return: 下载速度（KB/s）
        """
        try:
            start_time = time.time()

            # 下载一小块数据来测试速度
            response = self.session.get(url, stream=True, timeout=10)
            response.raise_for_status()

            downloaded = 0
            test_size = 1024 * 100  # 100KB

            for chunk in response.iter_content(chunk_size=1024):
                if chunk:
                    downloaded += len(chunk)
                    if downloaded >= test_size:
                        break

            end_time = time.time()
            duration = end_time - start_time

            if duration > 0:
                speed_kb_per_sec = (downloaded / 1024) / duration
                return speed_kb_per_sec

            return 0.0

        except Exception:
            return 0.0