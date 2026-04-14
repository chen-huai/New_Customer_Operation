# -*- coding: utf-8 -*-
"""New Customer Operation main entry."""

from __future__ import annotations

import logging
import os
import sys

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QWidget,
)

from New_Customer_Operate_Ui import Ui_MainWindow
from auto_updater import AutoUpdater, UI_AVAILABLE
from auto_updater.config_constants import CURRENT_VERSION
from config_manager import ConfigManager
from customer_mapper import CustomerMapper
from odm_api_client import OdmApiClient
from theme_manager_theme import ThemeManager
from word_form_parser import WordFormParser


class MainWindow(QMainWindow):
    """Main window."""

    def __init__(self, theme_manager: ThemeManager | None = None):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.logger = logging.getLogger(__name__)

        self.config_mgr = ConfigManager()
        self.updater: AutoUpdater | None = None
        self.word_parser = WordFormParser()
        self.customer_mapper = CustomerMapper()
        self.theme_manager = theme_manager

        self.current_parsed_form = None
        self.current_mapped_customer = None

        self.version_label: QLabel | None = None
        self.update_button: QPushButton | None = None
        self.toggle_theme_action: QAction | None = None

        self._setup_window()
        self._setup_theme_action()
        self._setup_auto_update()
        self._setup_status_bar()
        self._connect_signals()
        self._first_launch_check()
        self._load_default_import_path()

    def _setup_window(self) -> None:
        self.setWindowTitle("New Customer Operation")
        self.ui.statusbar.showMessage("就绪", 3000)

    def _setup_theme_action(self) -> None:
        if self.theme_manager is None or not self.theme_manager.is_available():
            return

        self.toggle_theme_action = QAction("Switch Theme", self)
        self.toggle_theme_action.setObjectName("actionSwitchTheme")
        self.ui.menuHelp.addSeparator()
        self.ui.menuHelp.addAction(self.toggle_theme_action)

    def _setup_status_bar(self) -> None:
        container = QWidget(self)
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        current_version = (
            self.updater.config.current_version
            if self.updater is not None
            else CURRENT_VERSION
        )
        self.version_label = QLabel(
            f"版本 v{current_version}",
            container,
        )
        self.version_label.setStyleSheet("color: #666;")

        self.update_button = QPushButton("Update", container)
        self.update_button.setFixedHeight(22)
        self.update_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.update_button.setStyleSheet(
            "QPushButton { padding: 2px 10px; }"
        )

        layout.addWidget(self.version_label)
        layout.addWidget(self.update_button)
        self.ui.statusbar.addPermanentWidget(container)

    def _setup_auto_update(self) -> None:
        """Initialize auto update integration."""
        try:
            if not UI_AVAILABLE:
                self.logger.warning("自动更新 UI 不可用，跳过初始化")
                self.updater = None
                return

            self.updater = AutoUpdater(parent=self)
            self.logger.info("自动更新功能初始化成功")
            self._startup_update_check()
        except Exception as exc:
            self.logger.error(f"自动更新器初始化失败: {exc}", exc_info=True)
            self.updater = None

    def _connect_signals(self) -> None:
        self.ui.actionExport.triggered.connect(self.export_config)
        self.ui.actionImport.triggered.connect(self.import_config)
        self.ui.actionUpdate.triggered.connect(self._on_check_update_clicked)
        self.ui.actionExit.triggered.connect(self.close)
        self.ui.actionAuthor.triggered.connect(self.show_author)
        self.ui.actionHelp.triggered.connect(self.show_help)
        self.ui.pushButton.clicked.connect(self.load_word_file)
        self.ui.pushButton_2.clicked.connect(self.submit_to_odm)
        if self.update_button is not None:
            self.update_button.clicked.connect(self._on_check_update_clicked)
        if self.toggle_theme_action is not None:
            self.toggle_theme_action.triggered.connect(self._toggle_theme)

    def _toggle_theme(self) -> None:
        if self.theme_manager is None or not self.theme_manager.toggle_theme():
            QMessageBox.information(
                self,
                "Theme",
                "No themes are available.",
            )
            return

        self.ui.statusbar.showMessage(
            f"Current theme: {self.theme_manager.current_theme}",
            3000,
        )

    def _on_check_update_clicked(self) -> None:
        """Handle manual update checks."""
        try:
            if self.updater is None:
                QMessageBox.warning(
                    self,
                    "更新功能不可用",
                    "自动更新功能未正确初始化，请检查配置或日志。",
                )
                return

            self.logger.info("用户手动触发更新检查")
            self.ui.statusbar.showMessage("正在检查更新...", 3000)
            self.updater.check_for_updates_with_ui(force_check=True)
        except Exception as exc:
            self.logger.error(f"检查更新失败: {exc}", exc_info=True)
            QMessageBox.critical(
                self,
                "更新检查失败",
                f"检查更新时出错：\n{exc}",
            )

    def _startup_update_check(self) -> None:
        """Run a delayed silent update check after startup."""
        if self.updater is None:
            return
        QTimer.singleShot(1000, self._perform_silent_check)

    def _perform_silent_check(self) -> None:
        """Perform a silent startup update check."""
        try:
            if self.updater is None:
                return

            has_update, remote_version, local_version, error = self.updater.check_for_updates(
                force_check=False,
                is_silent=True,
            )

            if has_update:
                self.logger.info(
                    "发现新版本: %s (当前版本: %s)",
                    remote_version,
                    local_version,
                )
                reply = QMessageBox.question(
                    self,
                    "发现新版本",
                    f"检测到新版本 {remote_version} (当前版本: {local_version})\n\n"
                    "是否立即查看更新详情？",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes,
                )
                if reply == QMessageBox.Yes:
                    self.updater.check_for_updates_with_ui(force_check=True)
            elif error:
                self.logger.debug(f"启动更新检查: {error}")
        except Exception as exc:
            self.logger.debug(f"静默更新检查异常: {exc}")

    def _first_launch_check(self) -> None:
        preview = self.config_mgr.get_sync_preview()
        if not preview["needs_sync"]:
            return

        reply = QMessageBox.question(
            self,
            "配置同步确认",
            preview["message"],
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if reply != QMessageBox.Yes:
            self.ui.statusbar.showMessage("已取消本次配置同步", 5000)
            return

        self.config_mgr.sync_config()
        QMessageBox.information(
            self,
            "配置同步完成",
            f"配置文件已更新：\n{self.config_mgr.config_path}",
        )

    def _load_default_import_path(self) -> None:
        try:
            config = self.config_mgr.read_config()
        except FileNotFoundError:
            return

        file_path = config.get("Files_Import_URL", "").strip()
        if file_path:
            self.ui.lineEdit.setText(file_path)

    def export_config(self) -> None:
        try:
            self.config_mgr.create_default_config()
            self.ui.statusbar.showMessage(
                f"配置文件已生成：{self.config_mgr.config_path}",
                5000,
            )
            QMessageBox.information(
                self,
                "导出成功",
                f"已生成配置文件：\n{self.config_mgr.config_path}",
            )
        except Exception as exc:
            QMessageBox.critical(
                self,
                "导出失败",
                f"生成配置文件失败：\n{exc}",
            )

    def import_config(self) -> None:
        try:
            config = self.config_mgr.read_config()
            lines = [f"{key}: {value}" for key, value in config.items()]
            self.ui.textBrowser.setText("\n".join(lines))
            self._load_default_import_path()
            self.ui.statusbar.showMessage(
                f"配置文件已导入：{self.config_mgr.config_path}",
                5000,
            )
        except FileNotFoundError:
            QMessageBox.warning(
                self,
                "导入失败",
                "配置文件不存在，请先执行 Export。\n\n"
                f"预期路径：{self.config_mgr.config_path}",
            )
        except Exception as exc:
            QMessageBox.critical(
                self,
                "导入失败",
                f"读取配置文件失败：\n{exc}",
            )

    def check_update(self) -> None:
        self._on_check_update_clicked()

    def load_word_file(self) -> None:
        current_path = self.ui.lineEdit.text().strip()
        if current_path and os.path.exists(current_path):
            start_dir = os.path.dirname(current_path)
        else:
            start_dir = os.getcwd()

        selected_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择客户 Word 文件",
            start_dir,
            "Word Files (*.docx *.doc)",
        )
        if not selected_path:
            return

        selected_path = os.path.abspath(selected_path)
        self.current_parsed_form = None
        self.current_mapped_customer = None

        self.ui.lineEdit.setText(selected_path)
        self.ui.textBrowser.setText(selected_path)
        self.config_mgr.set_config_value("Files_Import_URL", selected_path)
        self.ui.statusbar.showMessage(f"已选择文件：{selected_path}", 5000)

    def submit_to_odm(self) -> None:
        file_path = self.ui.lineEdit.text().strip()
        if not file_path:
            QMessageBox.warning(self, "缺少文件", "请先选择 Word 文件。")
            return

        if not os.path.exists(file_path):
            QMessageBox.warning(
                self,
                "文件不存在",
                f"请选择有效的 Word 文件：\n{file_path}",
            )
            return

        if not self.current_parsed_form or self.current_parsed_form["source_file"] != file_path:
            if not self._parse_file(file_path):
                return

        config = self.config_mgr.get_config()
        config_errors = self.config_mgr.validate_config(config)
        if config_errors:
            QMessageBox.warning(self, "配置不完整", "\n".join(config_errors))
            return

        applicant_payload, applicant_warnings = self.customer_mapper.build_applicant_create_payload(
            self.current_mapped_customer["customer_data"],
        )
        invoice_payload, payload_warnings = self.customer_mapper.build_invoice_create_payload(
            self.current_mapped_customer["customer_data"],
            config,
        )
        contact_payloads = self.customer_mapper.build_contact_payloads(
            self.current_mapped_customer["customer_data"],
            invoice_id=0,
        )

        warnings = list(self.current_mapped_customer["warnings"]) + applicant_warnings + payload_warnings
        preview_text = self.customer_mapper.format_preview(
            self.current_parsed_form,
            self.current_mapped_customer,
            applicant_payload=applicant_payload,
            invoice_payload=invoice_payload,
            contact_payloads=contact_payloads,
        )
        self.ui.textBrowser.setText(preview_text)

        blocking_warnings = applicant_warnings + payload_warnings
        if blocking_warnings:
            QMessageBox.warning(self, "提交前检查失败", "\n".join(blocking_warnings))
            return

        environment = config.get("Environment", "test").strip().lower()
        reply = QMessageBox.question(
            self,
            "确认提交到 ODM",
            "将执行以下操作：\n"
            f"1. 登录 {environment} 环境\n"
            "2. 创建 applicant\n"
            "3. 创建 customer / invoice\n"
            "4. 使用返回的 invoiceId 创建联系人\n\n"
            "是否继续？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            self.ui.statusbar.showMessage("已取消本次 ODM 提交", 5000)
            return

        try:
            client = OdmApiClient(config)
            client.login()
            client.create_applicant(applicant_payload)

            invoice_result = client.create_invoice(invoice_payload)
            invoice_id = invoice_result["id"]

            created_contacts = 0
            if contact_payloads:
                for contact in contact_payloads:
                    contact["invoiceId"] = invoice_id
                client.add_contacts(contact_payloads)
                created_contacts = len(contact_payloads)

            result_lines = [
                f"提交成功，环境：{environment}",
                f"applicant: {applicant_payload.get('name', '')}",
                f"invoiceId: {invoice_id}",
                f"联系人数量：{created_contacts}",
                "",
                "Invoice result:",
                str(invoice_result),
            ]
            if warnings:
                result_lines.extend(["", "Warnings:"] + warnings)

            self.ui.textBrowser.setText("\n".join(result_lines))
            self.ui.statusbar.showMessage(
                f"ODM 创建成功，invoiceId={invoice_id}",
                5000,
            )
            QMessageBox.information(
                self,
                "提交成功",
                "申请方与客户已创建\n"
                f"invoiceId={invoice_id}\n"
                f"联系人数量：{created_contacts}",
            )
        except Exception as exc:
            QMessageBox.critical(
                self,
                "提交失败",
                f"提交到 ODM 失败：\n{exc}",
            )

    def _parse_file(self, file_path: str) -> bool:
        try:
            parsed_form = self.word_parser.parse(file_path)
            mapped_customer = self.customer_mapper.map_to_customer_data(parsed_form)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "读取失败",
                f"读取 Word 文件失败：\n{exc}",
            )
            return False

        self.current_parsed_form = parsed_form
        self.current_mapped_customer = mapped_customer
        self.ui.lineEdit.setText(file_path)
        self.config_mgr.set_config_value("Files_Import_URL", file_path)
        self.ui.textBrowser.setText(
            self.customer_mapper.format_preview(parsed_form, mapped_customer)
        )
        self.ui.statusbar.showMessage(f"已读取文件：{file_path}", 5000)
        return True

    def show_author(self) -> None:
        QMessageBox.about(
            self,
            "Author",
            "New Customer Operation\n\nAuthor: chen, frank",
        )

    def show_help(self) -> None:
        QMessageBox.about(
            self,
            "Help",
            "使用说明：\n\n"
            "1. 获取文件：选择 .docx / .doc 文件，并回填绝对路径。\n"
            "2. 填入 ODM：读取 Word、解析字段，并提交 applicant / customer / contact。\n"
            "3. Export：生成配置文件。\n"
            "4. Import：读取配置文件。\n"
            "5. Update：检查新版本，确认后下载并执行更新。",
        )

    def closeEvent(self, event) -> None:  # noqa: N802
        try:
            if self.updater is not None:
                self.logger.info("正在清理自动更新器资源")
                self.updater.cleanup()
        finally:
            super().closeEvent(event)


def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    try:
        from auto_updater.auto_complete import auto_complete_update_if_needed

        def update_callback(success: bool, message: str) -> None:
            logger = logging.getLogger(__name__)
            if success:
                logger.info(f"后台自动完成更新成功: {message}")
            else:
                logger.info(f"后台自动完成更新状态: {message}")

        auto_complete_update_if_needed(update_callback)
    except Exception as exc:
        logging.getLogger(__name__).warning(f"自动完成更新检查失败: {exc}")

    app = QApplication(sys.argv)
    theme_manager = ThemeManager(app)
    theme_manager.set_default_theme()

    window = MainWindow(theme_manager=theme_manager)
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
