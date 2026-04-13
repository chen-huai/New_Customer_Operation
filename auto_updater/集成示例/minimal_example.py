#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AutoUpdater极简集成示例

这是最简单的集成方式，只需要3行代码即可添加完整的自动更新功能。
"""

import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QLabel, QVBoxLayout, QWidget
from auto_updater import AutoUpdater


class MinimalExampleWindow(QMainWindow):
    """极简集成示例窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("AutoUpdater 极简集成示例")
        self.setGeometry(100, 100, 600, 400)

        # 设置基本UI
        self.setup_ui()

        # 自动更新集成 - 仅需3行代码！
        self.auto_updater = AutoUpdater(self)
        self.auto_updater.setup_update_ui(self.menuBar(), "帮助(H)")

    def setup_ui(self):
        """设置基本界面"""
        central_widget = QWidget()
        layout = QVBoxLayout()

        # 添加说明文字
        info_label = QLabel("""
        <h2>AutoUpdater 极简集成示例</h2>

        <p>这个示例展示了如何用最少的代码集成自动更新功能。</p>

        <h3>集成代码：</h3>
        <code>
        self.auto_updater = AutoUpdater(self)<br>
        self.auto_updater.setup_update_ui(self.menuBar(), "帮助(H)")
        </code>

        <h3>功能特性：</h3>
        <ul>
        <li>✅ 自动版本检查</li>
        <li>✅ 一键更新下载</li>
        <li>✅ 自动备份和回滚</li>
        <li>✅ 进度显示和错误处理</li>
        <li>✅ 完整的用户界面</li>
        </ul>

        <p><b>使用说明：</b></p>
        <p>点击菜单栏的"帮助(H)" → "检查更新"来测试自动更新功能。</p>
        """)
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def closeEvent(self, event):
        """应用退出时清理资源"""
        if hasattr(self, 'auto_updater'):
            self.auto_updater.cleanup()
        event.accept()


def main():
    """主函数"""
    app = QApplication(sys.argv)

    # 创建主窗口
    window = MinimalExampleWindow()
    window.show()

    # 运行应用
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()