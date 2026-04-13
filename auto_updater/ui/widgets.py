# -*- coding: utf-8 -*-
"""
更新功能通用UI组件
提供可复用的UI组件和工具函数
"""
import logging
from typing import Optional, Union
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QProgressBar, QFrame)
from PyQt5.QtCore import Qt, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QPalette

from .resources import UpdateUIStyle

logger = logging.getLogger(__name__)

class UpdateStatusWidget(QWidget):
    """
    更新状态显示组件

    可嵌入到其他界面中，显示当前更新状态和提供快速操作
    """

    # 信号定义
    check_update_clicked = pyqtSignal()
    update_now_clicked = pyqtSignal(str)  # 参数：版本号

    def __init__(self, parent=None):
        """
        初始化更新状态组件

        Args:
            parent: 父组件
        """
        super().__init__(parent)
        self.current_version = "未知"
        self.remote_version = None
        self.has_update = False
        self.is_checking = False

        self._setup_ui()
        self._setup_style()

    def _setup_ui(self) -> None:
        """设置UI界面"""
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 5, 10, 5)

        # 主状态区域
        main_layout = QHBoxLayout()

        # 状态图标和文本
        self.status_icon_label = QLabel("📋")
        self.status_icon_label.setFixedSize(20, 20)
        self.status_icon_label.setAlignment(Qt.AlignCenter)

        self.status_text_label = QLabel("检查更新状态...")
        self.status_text_label.setWordWrap(True)

        # 版本信息
        self.version_label = QLabel("")
        self.version_label.setAlignment(Qt.AlignRight)

        main_layout.addWidget(self.status_icon_label)
        main_layout.addWidget(self.status_text_label, 1)
        main_layout.addWidget(self.version_label)

        # 按钮区域
        button_layout = QHBoxLayout()

        self.check_update_btn = QPushButton("检查更新")
        self.check_update_btn.clicked.connect(self._on_check_update_clicked)
        self.check_update_btn.setFixedSize(80, 25)

        self.update_now_btn = QPushButton("立即更新")
        self.update_now_btn.clicked.connect(self._on_update_now_clicked)
        self.update_now_btn.setFixedSize(80, 25)
        self.update_now_btn.hide()  # 默认隐藏

        button_layout.addStretch()
        button_layout.addWidget(self.check_update_btn)
        button_layout.addWidget(self.update_now_btn)

        # 进度条（通常隐藏）
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.hide()

        # 添加到主布局
        layout.addLayout(main_layout)
        layout.addLayout(button_layout)
        layout.addWidget(self.progress_bar)

        self.setLayout(layout)

    def _setup_style(self) -> None:
        """设置样式"""
        self.setStyleSheet(f"""
            UpdateStatusWidget {{
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 4px;
            }}
            QLabel {{
                color: #495057;
                font-size: 12px;
            }}
            QPushButton {{
                background-color: #007bff;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 2px 8px;
                font-size: 11px;
            }}
            QPushButton:hover {{
                background-color: #0056b3;
            }}
            QPushButton:disabled {{
                background-color: #6c757d;
            }}
        """)

    def set_current_version(self, version: str) -> None:
        """
        设置当前版本

        Args:
            version: 当前版本号
        """
        self.current_version = version
        self.version_label.setText(f"v{version}")
        self._update_display()

    def set_update_status(self, has_update: bool, remote_version: str = None) -> None:
        """
        设置更新状态

        Args:
            has_update: 是否有更新
            remote_version: 远程版本号
        """
        self.has_update = has_update
        self.remote_version = remote_version
        self.is_checking = False

        if has_update and remote_version:
            self.status_text_label.setText(f"发现新版本 v{remote_version}")
            self.status_icon_label.setText("🆕")
            self.update_now_btn.show()
            self.check_update_btn.setText("重新检查")
        else:
            self.status_text_label.setText("已是最新版本")
            self.status_icon_label.setText("✅")
            self.update_now_btn.hide()
            self.check_update_btn.setText("检查更新")

        self.check_update_btn.setEnabled(True)
        self._update_display()

    def set_checking_status(self) -> None:
        """设置检查状态"""
        self.is_checking = True
        self.status_text_label.setText("正在检查更新...")
        self.status_icon_label.setText("🔍")
        self.check_update_btn.setEnabled(False)
        self.update_now_btn.hide()
        self._update_display()

    def set_progress(self, value: int, status: str = "") -> None:
        """
        设置进度

        Args:
            value: 进度值 (0-100)
            status: 状态文本
        """
        # 数据验证和边界检查
        try:
            safe_value = int(value) if value is not None else 0
        except (ValueError, TypeError):
            safe_value = 0
            logger.warning(f"无效的进度值 {value}，已重置为0")

        # 限制进度值在合理范围内
        safe_value = max(0, min(100, safe_value))

        self.progress_bar.show()
        self.progress_bar.setValue(safe_value)

        if status:
            self.status_text_label.setText(status)

    def hide_progress(self) -> None:
        """隐藏进度条"""
        self.progress_bar.hide()

    def _update_display(self) -> None:
        """更新显示状态"""
        # 根据状态更新样式
        if self.is_checking:
            self.setStyleSheet(self.styleSheet() + """
                UpdateStatusWidget { background-color: #fff3cd; border-color: #ffeaa7; }
            """)
        elif self.has_update:
            self.setStyleSheet(self.styleSheet() + """
                UpdateStatusWidget { background-color: #d4edda; border-color: #c3e6cb; }
            """)
        else:
            self.setStyleSheet(self.styleSheet() + """
                UpdateStatusWidget { background-color: #f8f9fa; border-color: #dee2e6; }
            """)

    def _on_check_update_clicked(self) -> None:
        """检查更新按钮点击事件"""
        if not self.is_checking:
            self.check_update_clicked.emit()

    def _on_update_now_clicked(self) -> None:
        """立即更新按钮点击事件"""
        if self.remote_version:
            self.update_now_clicked.emit(self.remote_version)


class UpdateInfoWidget(QWidget):
    """
    更新信息显示组件

    显示详细的更新信息和操作选项
    """

    def __init__(self, parent=None):
        """
        初始化更新信息组件

        Args:
            parent: 父组件
        """
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self) -> None:
        """设置UI界面"""
        layout = QVBoxLayout()

        # 标题
        title_label = QLabel("更新信息")
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(14)
        title_label.setFont(title_font)

        # 分隔线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)

        # 信息显示区域
        self.info_label = QLabel("暂无更新信息")
        self.info_label.setWordWrap(True)
        self.info_label.setMinimumHeight(60)

        # 添加到布局
        layout.addWidget(title_label)
        layout.addWidget(line)
        layout.addWidget(self.info_label)
        layout.addStretch()

        self.setLayout(layout)

    def set_update_info(self, info: dict) -> None:
        """
        设置更新信息

        Args:
            info: 更新信息字典
        """
        info_text = f"""
        <b>当前版本:</b> {info.get('current_version', '未知')}<br>
        <b>最新版本:</b> {info.get('remote_version', '未知')}<br>
        <b>发布时间:</b> {info.get('release_date', '未知')}<br>
        <b>文件大小:</b> {info.get('file_size', '未知')}<br>
        <b>更新说明:</b> {info.get('release_notes', '暂无说明')}
        """
        self.info_label.setText(info_text)

    def set_error_info(self, error: str) -> None:
        """
        设置错误信息

        Args:
            error: 错误信息
        """
        error_text = f"<b>获取更新信息失败:</b><br>{error}"
        self.info_label.setText(error_text)


class QuickUpdateButton(QPushButton):
    """
    快速更新按钮

    简化的更新按钮，可嵌入到工具栏或状态栏
    """

    def __init__(self, parent=None):
        """
        初始化快速更新按钮

        Args:
            parent: 父组件
        """
        super().__init__(parent)
        self.has_update = False
        self.is_checking = False
        self._setup_button()

    def _setup_button(self) -> None:
        """设置按钮"""
        self.setText("检查更新")
        self.setFixedSize(80, 25)
        self.setStyleSheet(UpdateUIStyle.QUICK_BUTTON_STYLE)

    def set_has_update(self, has_update: bool, version: str = None) -> None:
        """
        设置是否有更新

        Args:
            has_update: 是否有更新
            version: 新版本号
        """
        self.has_update = has_update
        self.is_checking = False

        if has_update:
            self.setText(f"更新到 v{version[:8]}..." if version else "有更新")
            self.setStyleSheet(UpdateUIStyle.QUICK_BUTTON_UPDATE_STYLE)
        else:
            self.setText("检查更新")
            self.setStyleSheet(UpdateUIStyle.QUICK_BUTTON_STYLE)

        self.setEnabled(True)

    def set_checking(self) -> None:
        """设置为检查状态"""
        self.is_checking = True
        self.setText("检查中...")
        self.setEnabled(False)

    def set_error(self) -> None:
        """设置为错误状态"""
        self.is_checking = False
        self.setText("检查失败")
        self.setEnabled(True)