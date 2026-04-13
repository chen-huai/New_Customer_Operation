# -*- coding: utf-8 -*-
"""New Customer Operation main entry."""

from __future__ import annotations

import os
import sys

from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QMessageBox

from New_Customer_Operate_Ui import Ui_MainWindow
from auto_updater import AutoUpdater
from config_manager import ConfigManager
from customer_mapper import CustomerMapper
from odm_api_client import OdmApiClient
from word_form_parser import WordFormParser


class MainWindow(QMainWindow):
    """Main window."""

    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.config_mgr = ConfigManager()
        self.updater = AutoUpdater(parent=self)
        self.word_parser = WordFormParser()
        self.customer_mapper = CustomerMapper()

        self.current_parsed_form = None
        self.current_mapped_customer = None

        self._connect_signals()
        self._first_launch_check()
        self._load_default_import_path()

    def _connect_signals(self):
        self.ui.actionExport.triggered.connect(self.export_config)
        self.ui.actionImport.triggered.connect(self.import_config)
        self.ui.actionUpdate.triggered.connect(self.check_update)
        self.ui.actionExit.triggered.connect(self.close)
        self.ui.actionAuthor.triggered.connect(self.show_author)
        self.ui.actionHelp.triggered.connect(self.show_help)
        self.ui.pushButton.clicked.connect(self.load_word_file)
        self.ui.pushButton_2.clicked.connect(self.submit_to_odm)

    def _first_launch_check(self):
        preview = self.config_mgr.get_sync_preview()
        if not preview["needs_sync"]:
            return

        reply = QMessageBox.question(
            self,
            "配置更新确认",
            preview["message"],
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if reply != QMessageBox.Yes:
            self.ui.statusbar.showMessage("已取消本次配置更新", 5000)
            return

        self.config_mgr.sync_config()
        QMessageBox.information(
            self,
            "配置更新完成",
            f"配置文件已更新：\n{self.config_mgr.config_path}",
        )

    def _load_default_import_path(self):
        try:
            config = self.config_mgr.read_config()
        except FileNotFoundError:
            return

        file_path = config.get("Files_Import_URL", "").strip()
        if file_path:
            self.ui.lineEdit.setText(file_path)

    def export_config(self):
        try:
            self.config_mgr.create_default_config()
            self.ui.statusbar.showMessage(
                f"配置文件已重新生成：{self.config_mgr.config_path}",
                5000,
            )
            QMessageBox.information(
                self,
                "导出成功",
                f"已强制生成配置文件：\n{self.config_mgr.config_path}",
            )
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", f"生成配置文件失败：{exc}")

    def import_config(self):
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
            QMessageBox.critical(self, "导入失败", f"读取配置文件失败：{exc}")

    def check_update(self):
        try:
            self.updater.check_for_updates_with_ui(force_check=True)
        except Exception as exc:
            QMessageBox.critical(self, "更新检查失败", f"检查更新时出错：{exc}")

    def load_word_file(self):
        current_path = self.ui.lineEdit.text().strip()
        start_dir = os.path.dirname(current_path) if current_path and os.path.exists(current_path) else os.getcwd()
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
        self.config_mgr.set_config_value("Files_Import_URL", selected_path)
        self.ui.textBrowser.setText(selected_path)
        self.ui.statusbar.showMessage(f"已选择文件：{selected_path}", 5000)

    def submit_to_odm(self):
        file_path = self.ui.lineEdit.text().strip()
        if not file_path:
            QMessageBox.warning(self, "缺少文件", "请先选择或读取 Word 文件。")
            return
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "文件不存在", f"请选择有效的 Word 文件：\n{file_path}")
            return

        if not self.current_parsed_form or self.current_parsed_form["source_file"] != file_path:
            self._parse_file(file_path)
            if not self.current_parsed_form:
                return

        config = self.config_mgr.get_config()
        config_errors = self.config_mgr.validate_config(config)
        if config_errors:
            QMessageBox.warning(self, "配置不完整", "\n".join(config_errors))
            return

        invoice_payload, payload_warnings = self.customer_mapper.build_invoice_create_payload(
            self.current_mapped_customer["customer_data"],
            config,
        )
        contact_payloads = self.customer_mapper.build_contact_payloads(
            self.current_mapped_customer["customer_data"],
            invoice_id=0,
        )
        warnings = list(self.current_mapped_customer["warnings"]) + payload_warnings
        preview_text = self.customer_mapper.format_preview(
            self.current_parsed_form,
            self.current_mapped_customer,
            invoice_payload=invoice_payload,
            contact_payloads=contact_payloads,
        )
        self.ui.textBrowser.setText(preview_text)

        if payload_warnings:
            QMessageBox.warning(self, "创建前检查失败", "\n".join(payload_warnings))
            return

        environment = config.get("Environment", "test").strip().lower()
        reply = QMessageBox.question(
            self,
            "确认提交到 ODM",
            "将执行以下操作：\n"
            f"1. 登录 {environment} 环境\n"
            "2. 创建 customer/invoice\n"
            "3. 用返回的 invoiceId 创建联系人\n\n"
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
            invoice_result = client.create_invoice(invoice_payload)
            invoice_id = invoice_result["id"]

            created_contacts = 0
            if contact_payloads:
                for contact in contact_payloads:
                    contact["invoiceId"] = invoice_id
                client.add_contacts(contact_payloads)
                created_contacts = len(contact_payloads)

            result_lines = [
                f"提交成功，环境: {environment}",
                f"invoiceId: {invoice_id}",
                f"联系人数量: {created_contacts}",
                "",
                "Invoice result:",
                str(invoice_result),
            ]
            if warnings:
                result_lines.extend(["", "Warnings:", *warnings])

            self.ui.textBrowser.setText("\n".join(result_lines))
            self.ui.statusbar.showMessage(f"ODM 创建成功，invoiceId={invoice_id}", 5000)
            QMessageBox.information(
                self,
                "提交成功",
                f"客户已创建，invoiceId={invoice_id}\n联系人数量={created_contacts}",
            )
        except Exception as exc:
            QMessageBox.critical(self, "提交失败", f"提交到 ODM 失败：{exc}")

    def _parse_file(self, file_path: str):
        try:
            parsed_form = self.word_parser.parse(file_path)
            mapped_customer = self.customer_mapper.map_to_customer_data(parsed_form)
        except Exception as exc:
            QMessageBox.critical(self, "读取失败", f"读取 Word 文件失败：{exc}")
            return

        self.current_parsed_form = parsed_form
        self.current_mapped_customer = mapped_customer
        self.ui.lineEdit.setText(file_path)
        self.config_mgr.set_config_value("Files_Import_URL", file_path)
        self.ui.textBrowser.setText(
            self.customer_mapper.format_preview(parsed_form, mapped_customer)
        )
        self.ui.statusbar.showMessage(f"已读取文件：{file_path}", 5000)

    def show_author(self):
        QMessageBox.about(self, "Author", "New Customer Operation\n\nAuthor: chen, frank")

    def show_help(self):
        QMessageBox.about(
            self,
            "Help",
            "使用说明：\n\n"
            "1. 获取文件 - 选择 .docx/.doc 文件并回填绝对路径\n"
            "2. 填入ODM - 读取 Word、解析字段、登录接口并创建 customer + contact\n"
            "3. Export - 生成配置文件\n"
            "4. Import - 读取配置文件\n"
            "5. Update - 检查软件更新",
        )


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
