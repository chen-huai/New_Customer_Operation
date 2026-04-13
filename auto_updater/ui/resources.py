# -*- coding: utf-8 -*-
"""
更新功能UI资源和常量
集中管理UI文本、样式和配置信息
"""
from PyQt5.QtCore import QSize

class UpdateUIText:
    """
    更新功能UI文本常量
    集中管理所有显示文本，便于国际化和维护
    """

    # 通用文本
    APP_NAME = "PDF重命名工具"
    VERSION_PREFIX = "版本 "
    ERROR_TITLE = "错误"
    INFO_TITLE = "信息"
    WARNING_TITLE = "警告"

    # 菜单相关
    HELP_MENU_TITLE = "帮助(H)"
    UPDATE_MENU_TEXT = "Update"
    UPDATE_MENU_TOOLTIP = "检查软件更新"
    VERSION_MENU_TOOLTIP = "显示关于对话框"
    VERSION_MENU_TEXT = "关于"

    # 检查更新
    CHECK_UPDATE_TITLE = "检查更新"
    CHECKING_UPDATE_MESSAGE = "正在检查更新..."
    CHECK_UPDATE_FAILED_TITLE = "检查更新失败"
    CHECK_UPDATE_FAILED_MESSAGE = "检查更新时发生错误"
    CHECK_UPDATE_ERROR_MESSAGE = "检查更新时发生异常"
    CHECK_UPDATE_UNAVAILABLE_MESSAGE = "自动更新功能未正确初始化"

    # 版本状态
    LATEST_VERSION_MESSAGE = "您的软件已是最新版本！"
    LATEST_VERSION_MESSAGE_SIMPLE = "✅ 已是最新版本"
    NEW_VERSION_FOUND_TITLE = "发现新版本"
    NEW_VERSION_FOUND_MESSAGE = "发现新版本 {remote_version}！\n当前版本: {local_version}\n\n是否立即下载更新？"
    NEW_VERSION_FOUND_MESSAGE_SIMPLE = "🆕 发现新版本:"

    # 更新过程
    UPDATING_TITLE = "正在更新"
    UPDATING_TO_VERSION_MESSAGE = "正在更新到版本"
    PREPARING_UPDATE_MESSAGE = "准备更新..."
    DOWNLOADING_UPDATE_MESSAGE = "正在下载更新文件..."
    INSTALLING_UPDATE_MESSAGE = "正在安装更新..."
    DOWNLOAD_FAILED_MESSAGE = "下载失败"
    INSTALL_FAILED_MESSAGE = "安装失败"
    UPDATE_PROCESS_ERROR_MESSAGE = "更新过程异常"
    START_UPDATE_ERROR_MESSAGE = "启动更新过程时发生错误"

    # 进度显示
    FILE_DOWNLOADED_TO_MESSAGE = "文件已下载到："
    RETRYING_MESSAGE = "等待重试，"
    UPDATE_COMPLETE_MESSAGE = "更新完成！应用程序将重启..."
    UPDATE_COMPLETE_TITLE = "更新完成"
    UPDATE_FAILED_MESSAGE = "更新失败"
    UPDATE_FAILED_TITLE = "更新失败"
    UPDATE_CANCELLED_MESSAGE = "暂时不更新"

    # 对话框
    ABOUT_DIALOG_TITLE = "关于 PDF重命名工具"
    LOADING_VERSION_MESSAGE = "正在加载版本信息..."
    CHECKING_UPDATE_STATUS_MESSAGE = "正在检查更新状态..."
    UPDATE_UNAVAILABLE_MESSAGE = "自动更新功能不可用"
    GET_VERSION_ERROR_MESSAGE = "获取版本信息失败"
    CHECK_UPDATE_STATUS_ERROR_MESSAGE = "检查更新状态失败"
    CANNOT_CHECK_UPDATE_STATUS_MESSAGE = "无法检查更新状态"
    GET_RELEASE_NOTES_ERROR_MESSAGE = "获取更新日志失败"
    GET_RELEASE_NOTES_UNAVAILABLE_MESSAGE = "无法获取更新日志：更新功能不可用"
    VIEW_RELEASE_NOTES_ERROR_MESSAGE = "查看更新日志失败"

    # 按钮
    CHECK_UPDATE_BUTTON_TEXT = "检查更新"
    VIEW_RELEASE_NOTES_BUTTON_TEXT = "查看更新日志"
    CLOSE_BUTTON_TEXT = "关闭"
    CANCEL_BUTTON_TEXT = "取消"

    # 发布说明
    RELEASE_NOTES_TITLE = "更新日志"
    RELEASE_NOTES_MESSAGE = "版本"

    # GitHub
    GITHUB_LINK_TEXT = "GitHub: chen-huai/Temu_PDF_Rename_APP"
    GITHUB_URL = "https://github.com/chen-huai/Temu_PDF_Rename_APP"
    OPEN_GITHUB_ERROR_MESSAGE = "无法打开GitHub页面"

    # 重启
    RESTART_FAILED_TITLE = "重启失败"
    RESTART_FAILED_MESSAGE = "重启应用程序失败"

    # 状态栏
    UPDATE_UNAVAILABLE_MESSAGE = "自动更新功能不可用"
    SHOW_ABOUT_ERROR_MESSAGE = "显示关于对话框时发生错误"

    # 组件文本
    NO_UPDATE_INFO = "暂无更新信息"
    GET_UPDATE_INFO_FAILED = "获取更新信息失败"

    # 快速按钮
    QUICK_UPDATE_TEXT = "检查更新"
    QUICK_UPDATE_TEXT_UPDATE = "有更新"
    QUICK_UPDATE_TEXT_CHECKING = "检查中..."
    QUICK_UPDATE_TEXT_ERROR = "检查失败"


class UpdateUIStyle:
    """
    更新功能UI样式常量
    定义组件尺寸、颜色和样式
    """

    # 对话框尺寸
    PROGRESS_DIALOG_SIZE = QSize(400, 150)
    ABOUT_DIALOG_SIZE = QSize(450, 400)

    # 标签样式
    TITLE_LABEL_STYLE = "font-size: 18px; font-weight: bold;"
    BUILD_INFO_STYLE = "color: gray; font-size: 10px;"
    PROGRESS_LABEL_STYLE = "color: #666; font-size: 11px; margin-top: 5px;"

    # 状态样式
    STATUS_UPDATE_STYLE = "color: #28a745; font-weight: bold;"  # 绿色
    STATUS_CURRENT_STYLE = "color: #007bff;"  # 蓝色
    STATUS_ERROR_STYLE = "color: #dc3545;"  # 红色

    # 链接样式
    LINK_STYLE = "color: #007bff; text-decoration: underline;"

    # 按钮样式
    QUICK_BUTTON_STYLE = """
        QPushButton {
            background-color: #6c757d;
            color: white;
            border: none;
            border-radius: 3px;
            padding: 2px 8px;
            font-size: 11px;
        }
        QPushButton:hover {
            background-color: #5a6268;
        }
        QPushButton:disabled {
            background-color: #adb5bd;
        }
    """

    QUICK_BUTTON_UPDATE_STYLE = """
        QPushButton {
            background-color: #28a745;
            color: white;
            border: none;
            border-radius: 3px;
            padding: 2px 8px;
            font-size: 11px;
        }
        QPushButton:hover {
            background-color: #218838;
        }
    """

    # 进度条样式
    PROGRESS_BAR_STYLE = """
        QProgressBar {
            border: 1px solid #ccc;
            border-radius: 3px;
            text-align: center;
            background-color: #f0f0f0;
        }
        QProgressBar::chunk {
            background-color: #007bff;
            border-radius: 2px;
        }
    """

    # 状态指示器颜色
    STATUS_COLORS = {
        'success': '#28a745',  # 绿色
        'warning': '#ffc107',  # 黄色
        'error': '#dc3545',    # 红色
        'info': '#007bff',     # 蓝色
        'secondary': '#6c757d' # 灰色
    }

    # 动画持续时间（毫秒）
    ANIMATION_DURATION = 300

    # 刷新间隔（秒）
    AUTO_REFRESH_INTERVAL = 30

    # 超时设置
    NETWORK_TIMEOUT = 30  # 秒
    UPDATE_TIMEOUT = 300  # 秒

    # 图标
    ICONS = {
        'check': '🔍',
        'success': '✅',
        'update': '🆕',
        'error': '❌',
        'warning': '⚠️',
        'info': 'ℹ️',
        'download': '⬇️',
        'install': '🔧',
        'restart': '🔄'
    }


class UpdateUIConfig:
    """
    更新功能UI配置
    定义UI行为和交互配置
    """

    # 自动更新配置
    AUTO_CHECK_ON_STARTUP = True
    AUTO_CHECK_INTERVAL = 86400  # 24小时（秒）
    STARTUP_CHECK_DELAY = 3000  # 启动检查延迟（毫秒）

    # 下载配置
    MAX_DOWNLOAD_RETRIES = 3
    RETRY_DELAY = 5  # 秒
    CHUNK_SIZE = 8192  # 字节

    # UI配置
    SHOW_DETAILED_PROGRESS = True
    SHOW_DOWNLOAD_SPEED = True
    SHOW_ESTIMATED_TIME = True

    # 对话框配置
    AUTO_CLOSE_UPDATE_DIALOG = True
    UPDATE_DIALOG_CLOSE_DELAY = 2000  # 毫秒
    SHOW_CONFIRMATION_DIALOGS = True

    # 日志配置
    ENABLE_UI_LOGGING = True
    LOG_LEVEL = 'INFO'

    # 性能配置
    ENABLE_ANIMATIONS = True
    SMOOTH_PROGRESS = True
    THROTTLE_UI_UPDATES = True

    # 网络配置
    CONCURRENT_DOWNLOADS = 1
    TIMEOUT = 30  # 秒

    # 备份配置
    AUTO_BACKUP_BEFORE_UPDATE = True
    MAX_BACKUP_FILES = 5

    # 安全配置
    VERIFY_DOWNLOAD_INTEGRITY = True
    REQUIRE_USER_CONFIRMATION = True