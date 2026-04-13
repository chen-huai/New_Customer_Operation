# -*- coding: utf-8 -*-
"""
通用重试工具类
提供可复用的重试机制和错误判断逻辑
"""

import time
from typing import Callable, Type, Union, Tuple, Optional, Any
from abc import ABC, abstractmethod

class RetryableError(Exception):
    """可重试的错误基类"""
    pass

class NonRetryableError(Exception):
    """不可重试的错误基类"""
    pass

class RetryStrategy(ABC):
    """重试策略抽象基类"""

    @abstractmethod
    def should_retry(self, exception: Exception, attempt_count: int) -> bool:
        """
        判断是否应该重试
        :param exception: 异常对象
        :param attempt_count: 当前尝试次数
        :return: 是否应该重试
        """
        pass

    @abstractmethod
    def get_delay(self, attempt_count: int) -> float:
        """
        获取重试延迟时间
        :param attempt_count: 当前尝试次数
        :return: 延迟时间（秒）
        """
        pass

class DefaultRetryStrategy(RetryStrategy):
    """默认重试策略"""

    def __init__(self, max_retries: int = 3, base_delay: float = 2.0, max_delay: float = 30.0):
        """
        初始化重试策略
        :param max_retries: 最大重试次数
        :param base_delay: 基础延迟时间
        :param max_delay: 最大延迟时间
        """
        self.max_retries = max_retries
        self.base_delay = base_delay
        self.max_delay = max_delay

    def should_retry(self, exception: Exception, attempt_count: int) -> bool:
        """
        判断是否应该重试
        """
        if attempt_count >= self.max_retries:
            return False

        # 如果是不可重试错误，直接返回False
        if isinstance(exception, NonRetryableError):
            return False

        # 如果是可重试错误，直接返回True
        if isinstance(exception, RetryableError):
            return True

        # 根据异常类型和消息进行判断
        error_message = str(exception).lower()
        exception_type = type(exception).__name__.lower()

        # 可重试的错误类型
        retryable_patterns = [
            # 网络相关
            "timeout", "connection", "network",
            # 服务器错误
            "http error 5", "502", "503", "504", "500",
            # 临时错误
            "temporary", "retry", "rate limit"
        ]

        # 不可重试的错误类型
        non_retryable_patterns = [
            # 认证错误
            "unauthorized", "401", "forbidden", "403",
            # 资源不存在
            "not found", "404",
            # 客户端错误
            "bad request", "400", "422",
            # SSL/DNS错误
            "dns", "ssl", "certificate",
            # 文件系统错误
            "permission", "access denied", "file not found"
        ]

        # 检查不可重试模式
        for pattern in non_retryable_patterns:
            if pattern in error_message or pattern in exception_type:
                return False

        # 检查可重试模式
        for pattern in retryable_patterns:
            if pattern in error_message or pattern in exception_type:
                return True

        # 默认不重试未知错误
        return False

    def get_delay(self, attempt_count: int) -> float:
        """
        计算重试延迟时间（指数退避）
        """
        if attempt_count <= 0:
            return 0

        delay = self.base_delay * (2 ** (attempt_count - 1))
        return min(delay, self.max_delay)

class NetworkRetryStrategy(DefaultRetryStrategy):
    """网络专用重试策略"""

    def should_retry(self, exception: Exception, attempt_count: int) -> bool:
        """
        网络错误的特殊重试判断
        """
        if attempt_count >= self.max_retries:
            return False

        error_message = str(exception).lower()

        # 对于频率限制，延长等待时间
        if "rate limit" in error_message or "429" in error_message:
            return True

        # 调用父类方法
        return super().should_retry(exception, attempt_count)

    def get_delay(self, attempt_count: int) -> float:
        """
        网络重试的特殊延迟计算
        """
        # 频率限制特殊处理
        if attempt_count >= 2:
            return 60  # 第三次重试等待1分钟

        return super().get_delay(attempt_count)

class RetryExecutor:
    """重试执行器"""

    def __init__(self, strategy: RetryStrategy = None):
        """
        初始化重试执行器
        :param strategy: 重试策略，默认使用DefaultRetryStrategy
        """
        self.strategy = strategy or DefaultRetryStrategy()

    def execute(self, func: Callable, *args, **kwargs) -> Any:
        """
        执行带重试的函数调用
        :param func: 要执行的函数
        :param args: 函数参数
        :param kwargs: 函数关键字参数
        :return: 函数执行结果
        """
        last_exception = None

        for attempt in range(self.strategy.max_retries + 1):  # +1 包含首次尝试
            try:
                if attempt > 0:
                    delay = self.strategy.get_delay(attempt)
                    if delay > 0:
                        print(f"重试第 {attempt} 次，等待 {delay:.1f} 秒...")
                        time.sleep(delay)

                print(f"尝试第 {attempt + 1} 次执行...")
                result = func(*args, **kwargs)

                if attempt > 0:
                    print(f"重试成功，共尝试 {attempt + 1} 次")

                return result

            except Exception as e:
                last_exception = e
                should_retry = self.strategy.should_retry(e, attempt)

                if not should_retry:
                    print(f"错误不适合重试: {e}")
                    break

                print(f"执行失败（可重试）: {e}")

        # 所有重试都失败
        if last_exception:
            raise last_exception
        else:
            raise Exception("重试执行失败，原因未知")

# 便捷函数
def retry(max_retries: int = 3, base_delay: float = 2.0, max_delay: float = 30.0):
    """
    重试装饰器
    :param max_retries: 最大重试次数
    :param base_delay: 基础延迟时间
    :param max_delay: 最大延迟时间
    """
    def decorator(func):
        def wrapper(*args, **kwargs):
            strategy = DefaultRetryStrategy(max_retries, base_delay, max_delay)
            executor = RetryExecutor(strategy)
            return executor.execute(func, *args, **kwargs)
        return wrapper
    return decorator

def network_retry(max_retries: int = 3, base_delay: float = 2.0):
    """
    网络重试装饰器
    :param max_retries: 最大重试次数
    :param base_delay: 基础延迟时间
    """
    def decorator(func):
        def wrapper(*args, **kwargs):
            strategy = NetworkRetryStrategy(max_retries, base_delay)
            executor = RetryExecutor(strategy)
            return executor.execute(func, *args, **kwargs)
        return wrapper
    return decorator