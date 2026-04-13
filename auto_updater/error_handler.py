# -*- coding: utf-8 -*-
"""
统一错误处理模块
提供用户友好的错误信息和异常处理
"""

import traceback
import logging
from typing import Optional, Tuple
from enum import Enum

# 配置日志记录器
logger = logging.getLogger(__name__)

class ErrorType(Enum):
    """错误类型枚举"""
    NETWORK_ERROR = "网络连接错误"
    DOWNLOAD_ERROR = "文件下载错误"
    VERSION_ERROR = "版本检查错误"
    BACKUP_ERROR = "备份操作错误"
    UPDATE_ERROR = "更新执行错误"
    PERMISSION_ERROR = "权限不足"
    FILE_ERROR = "文件操作错误"
    CONFIG_ERROR = "配置错误"
    ENVIRONMENT_ERROR = "环境检测错误"
    SOURCE_FILE_ERROR = "源文件错误"
    DEVELOPMENT_ERROR = "开发环境错误"
    UNKNOWN_ERROR = "未知错误"

class UserFriendlyError:
    """用户友好的错误信息"""

    @staticmethod
    def get_user_message(error_type: ErrorType, technical_details: str = "") -> str:
        """
        获取用户友好的错误信息
        :param error_type: 错误类型
        :param technical_details: 技术细节（用于日志）
        :return: 用户友好的错误信息
        """
        messages = {
            ErrorType.NETWORK_ERROR: "无法连接到更新服务器。建议：\n1. 检查网络连接是否正常\n2. 尝试访问其他网站确认网络状态\n3. 如果是公司网络，请联系网络管理员\n4. 稍后重试或更换网络环境",
            ErrorType.DOWNLOAD_ERROR: "下载更新文件失败。建议：\n1. 检查网络连接稳定性\n2. 确保磁盘空间充足\n3. 关闭其他下载程序\n4. 稍后重试或联系技术支持",
            ErrorType.VERSION_ERROR: "版本检查失败。建议：\n1. 检查网络连接\n2. 确认软件版本信息正确\n3. 重启程序后重试\n4. 如持续失败，请联系技术支持",
            ErrorType.BACKUP_ERROR: "创建备份失败。建议：\n1. 确保磁盘空间充足\n2. 检查文件权限设置\n3. 关闭其他可能占用文件的程序\n4. 继续更新将无法回滚，请谨慎操作",
            ErrorType.UPDATE_ERROR: "更新过程中发生错误。建议：\n1. 重新下载完整程序\n2. 手动替换程序文件\n3. 检查系统完整性\n4. 联系技术支持获取帮助",
            ErrorType.PERMISSION_ERROR: "权限不足。建议：\n1. 右键程序选择'以管理员身份运行'\n2. 联系IT管理员获取必要权限\n3. 确保程序安装在可写目录",
            ErrorType.FILE_ERROR: "文件操作失败。建议：\n1. 确保磁盘空间充足（至少需要100MB）\n2. 关闭其他可能占用文件的程序\n3. 检查文件是否为只读属性\n4. 重启计算机后重试",
            ErrorType.CONFIG_ERROR: "配置文件错误。建议：\n1. 重新安装程序\n2. 检查程序文件完整性\n3. 清除临时文件后重试\n4. 联系技术支持",
            ErrorType.ENVIRONMENT_ERROR: "运行环境检测失败。建议：\n1. 确保Windows系统版本兼容\n2. 安装必要的运行库\n3. 更新系统组件\n4. 在兼容模式下运行程序",
            ErrorType.SOURCE_FILE_ERROR: "源文件不存在或无法访问。建议：\n1. 重新下载完整程序包\n2. 检查程序文件完整性\n3. 从官方渠道获取最新版本\n4. 进行病毒扫描确认文件安全",
            ErrorType.DEVELOPMENT_ERROR: "开发环境更新失败。建议：\n1. 检查源文件完整性\n2. 确认开发工具链正常\n3. 更新依赖项到兼容版本\n4. 检查版本控制系统状态",
            ErrorType.UNKNOWN_ERROR: "发生未知错误。建议：\n1. 记录错误详情和时间\n2. 重启程序后重试\n3. 更新到最新版本\n4. 如持续发生，请联系技术支持并提供错误日志"
        }

        return messages.get(error_type, messages[ErrorType.UNKNOWN_ERROR])

    @staticmethod
    def classify_error(exception: Exception) -> ErrorType:
        """
        根据异常类型分类错误
        :param exception: 异常对象
        :return: 错误类型
        """
        error_message = str(exception).lower()
        exception_type = type(exception).__name__.lower()

        # 优先检查具体的错误类型
        if "networkerror" in exception_type or "connectionerror" in exception_type:
            return ErrorType.NETWORK_ERROR
        elif "downloaderror" in exception_type:
            return ErrorType.DOWNLOAD_ERROR
        elif "versioncheckerror" in exception_type:
            return ErrorType.VERSION_ERROR

        # 然后检查错误消息内容
        if any(keyword in error_message for keyword in [
            "dns解析失败", "name resolution failed", "getaddrinfo failed",
            "网络连接失败", "connection failed", "connection refused",
            "ssl证书验证失败", "certificate verification failed"
        ]):
            return ErrorType.NETWORK_ERROR

        elif any(keyword in error_message for keyword in [
            "下载超时", "download timeout", "请求超时",
            "下载文件不存在", "下载失败", "访问被拒绝"
        ]):
            return ErrorType.DOWNLOAD_ERROR

        elif any(keyword in error_message for keyword in [
            "获取最新release失败", "版本检查失败", "仓库或release不存在",
            "release信息", "release不存在", "api请求频率限制"
        ]):
            return ErrorType.VERSION_ERROR

        elif "backup" in error_message or "备份" in error_message:
            return ErrorType.BACKUP_ERROR

        elif any(keyword in error_message for keyword in [
            "permission", "access denied", "权限不足", "拒绝访问"
        ]):
            return ErrorType.PERMISSION_ERROR

        elif any(keyword in error_message for keyword in [
            "源文件不存在", "source file", "file not found",
            "文件不存在", "找不到文件"
        ]):
            return ErrorType.SOURCE_FILE_ERROR

        elif any(keyword in error_message for keyword in [
            "开发环境", "development", "源码", "source code"
        ]):
            return ErrorType.DEVELOPMENT_ERROR

        elif any(keyword in error_message for keyword in [
            "环境检测", "environment", "运行环境"
        ]):
            return ErrorType.ENVIRONMENT_ERROR

        elif any(keyword in error_message for keyword in [
            "文件操作", "file operation", "disk space",
            "磁盘空间", "文件占用", "file in use"
        ]):
            return ErrorType.FILE_ERROR

        elif any(keyword in error_message for keyword in [
            "config", "配置", "json解析", "配置文件"
        ]):
            return ErrorType.CONFIG_ERROR

        else:
            return ErrorType.UNKNOWN_ERROR

class ErrorHandler:
    """错误处理器"""

    @staticmethod
    def handle_exception(exception: Exception, context: str = "") -> Tuple[str, ErrorType]:
        """
        处理异常并返回用户友好的错误信息
        :param exception: 异常对象
        :param context: 上下文信息
        :return: (用户友好的错误信息, 错误类型)
        """
        try:
            # 分类错误
            error_type = UserFriendlyError.classify_error(exception)

            # 获取用户友好的错误信息
            user_message = UserFriendlyError.get_user_message(error_type)

            # 记录技术详情到日志
            technical_details = f"Context: {context}\nException: {exception}\nTraceback: {traceback.format_exc()}"
            logger.error(f"技术错误详情:\n{technical_details}")

            return user_message, error_type

        except Exception as e:
            # 如果错误处理本身出错，返回基本错误信息
            logger.error(f"错误处理异常: {str(e)}\n{traceback.format_exc()}")
            return f"处理错误时发生异常: {str(e)}", ErrorType.UNKNOWN_ERROR

    @staticmethod
    def log_error(error_type: ErrorType, message: str, technical_details: str = ""):
        """
        记录错误日志
        :param error_type: 错误类型
        :param message: 错误消息
        :param technical_details: 技术细节
        """
        log_message = f"[{error_type.value}] {message}"
        if technical_details:
            log_message += f"\n技术详情: {technical_details}"

        logger.error(log_message)

    @staticmethod
    def log_info(message: str):
        """记录信息日志"""
        logger.info(message)

    @staticmethod
    def log_warning(message: str):
        """记录警告日志"""
        logger.warning(message)