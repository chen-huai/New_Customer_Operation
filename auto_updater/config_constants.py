# -*- coding: utf-8 -*-
"""
自动更新器配置常量
将JSON配置信息转换为Python常量，消除外部文件依赖
"""

# 应用配置
SOFTWARE_ID: str = "New_Customer_Operate_tool"  # 软件唯一标识，用于多软件状态隔离
APP_NAME: str = "New_Customer_Operate"
APP_EXECUTABLE: str = "New_Customer_Operate.exe"

# GitHub仓库配置
GITHUB_OWNER: str = "chen-huai"
GITHUB_REPO: str = "New_Customer_Operate"
GITHUB_API_BASE: str = "https://api.github.com"

# 版本配置
CURRENT_VERSION: str = "1.0.1"
UPDATE_CHECK_INTERVAL_DAYS: int = 30
AUTO_CHECK_ENABLED: bool = True

# 更新配置
MAX_BACKUP_COUNT: int = 3
DOWNLOAD_TIMEOUT: int = 600  # 下载文件超时（10分钟）
MAX_RETRIES: int = 3
AUTO_RESTART: bool = True

# 网络配置类
class NetworkConfig:
    """统一的网络配置管理"""

    # 超时配置
    TIMEOUTS = {
        'check': 15,        # 检查更新超时（秒）
        'download': 600,    # 下载文件超时（10分钟）
        'connection': 10,   # 连接超时（秒）
        'dns': 5           # DNS解析超时（秒）
    }

    # 重试配置
    RETRY = {
        'max_retries': 3,       # 最大重试次数
        'base_delay': 2,        # 基础延迟时间（秒）
        'max_delay': 30,        # 最大延迟时间（秒）
        'api_retry_delay': 60   # API频率限制延迟（秒）
    }

    # 请求头配置
    HEADERS = {
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "PDF-Rename-Tool-Updater/2.0.0",
        "Connection": "keep-alive",
        "Accept-Encoding": "gzip, deflate"
    }

    # 连接池配置
    CONNECTION_POOL = {
        'pool_connections': 10,  # 连接池大小
        'pool_maxsize': 20,      # 每个连接池的最大连接数
        'max_retries': 2,        # urllib3重试次数
        'backoff_factor': 2      # 重试间隔因子
    }

# 保持向后兼容的常量
CHECK_TIMEOUT: int = NetworkConfig.TIMEOUTS['check']
CONNECTION_TIMEOUT: int = NetworkConfig.TIMEOUTS['connection']
RETRY_DELAY: int = NetworkConfig.RETRY['base_delay']
DOWNLOAD_TIMEOUT: int = NetworkConfig.TIMEOUTS['download']
REQUEST_HEADERS: dict = NetworkConfig.HEADERS
MAX_RETRIES: int = NetworkConfig.RETRY['max_retries']

# 便利常量（兼容现有API）
GITHUB_REPO_PATH: str = f"{GITHUB_OWNER}/{GITHUB_REPO}"
GITHUB_RELEASES_URL: str = f"{GITHUB_API_BASE}/repos/{GITHUB_REPO_PATH}/releases"
GITHUB_LATEST_RELEASE_URL: str = f"{GITHUB_RELEASES_URL}/latest"

# 默认配置字典（保持JSON格式兼容）
DEFAULT_CONFIG: dict = {
    "app": {
        "software_id": SOFTWARE_ID,
        "name": APP_NAME,
        "executable": APP_EXECUTABLE
    },
    "repository": {
        "owner": GITHUB_OWNER,
        "repo": GITHUB_REPO,
        "api_base": GITHUB_API_BASE
    },
    "version": {
        "current": CURRENT_VERSION,
        "check_interval_days": UPDATE_CHECK_INTERVAL_DAYS,
        "auto_check_enabled": AUTO_CHECK_ENABLED
    },
    "update": {
        "backup_count": MAX_BACKUP_COUNT,
        "download_timeout": DOWNLOAD_TIMEOUT,
        "max_retries": MAX_RETRIES,
        "auto_restart": AUTO_RESTART
    },
    "network": {
        "request_headers": REQUEST_HEADERS
    }
}

# 版本信息验证
def validate_version_format(version_str: str) -> bool:
    """验证版本号格式是否有效"""
    try:
        from packaging import version as pkg_version
        pkg_version.parse(version_str)
        return True
    except Exception:
        return False

# 配置完整性验证
def validate_config() -> bool:
    """验证配置信息的完整性"""
    try:
        # 验证必要常量
        required_constants = [
            SOFTWARE_ID, APP_NAME, APP_EXECUTABLE, GITHUB_OWNER, GITHUB_REPO,
            GITHUB_API_BASE, CURRENT_VERSION, REQUEST_HEADERS
        ]

        for const in required_constants:
            if not const:
                return False

        # 验证版本号格式
        if not validate_version_format(CURRENT_VERSION):
            return False

        # 验证URL格式
        if not GITHUB_API_BASE.startswith("https://"):
            return False

        return True
    except Exception:
        return False

# 在模块加载时验证配置
if not validate_config():
    raise ValueError("配置信息验证失败，请检查config_constants.py中的配置")