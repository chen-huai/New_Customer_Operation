# -*- coding: utf-8 -*-
"""Map parsed Word form data into normalized customer fields and ODM payloads."""

from __future__ import annotations

import json
import re
from collections import OrderedDict


class CustomerMapper:
    """Transform parsed tables into normalized data and ODM payloads."""

    INVOICE_TYPE_NO_NEED = 0
    INVOICE_TYPE_SPECIAL = 1
    INVOICE_TYPE_INVOICE = 2
    INVOICE_TYPE_NORMAL = 3

    SITE_ID_MAP = {
        "0483 BJ": 1,
        "0484 GZ": 2,
        "0487 NB": 3,
        "0486 XM": 4,
    }

    def map_to_customer_data(self, parsed_form: dict) -> dict:
        tables = parsed_form["tables"]
        warnings = []
        customer_data = OrderedDict()

        customer_data["source_file"] = parsed_form["source_name"]
        customer_data["branch_codes"] = self._all_checked_labels(tables[0]) if len(tables) > 0 else []
        customer_data["branch_code"] = customer_data["branch_codes"][0] if customer_data["branch_codes"] else ""
        customer_data["customer_category"] = self._first_checked_label(tables[1]) if len(tables) > 1 else ""

        if not customer_data["branch_codes"]:
            warnings.append("未检测到 branch/site 勾选项。")
        if not customer_data["customer_category"]:
            warnings.append("Customer category 未检测到勾选项。")

        if len(tables) > 1:
            rows = tables[1]["rows"]
            customer_data["new_order_value"] = self._get_row_value(rows, "New order value", 1)
            customer_data["currency"] = self._get_row_value(rows, "New order value", 2)
            customer_data["industry_code"] = self._get_row_value(rows, "Industry Code", 1)
            customer_data["customer_code"] = self._get_row_value(rows, "Industry Code", 3)

        if len(tables) > 2:
            rows = tables[2]["rows"]
            customer_data["customer_name_en"] = self._clean_company_name(
                self._get_first_matching_value(
                    rows,
                    ["Customer Name (EN)", "Customer Name (EN)客户公司英文名称"],
                )
            )
            customer_data["customer_name_cn"] = self._clean_company_name(
                self._get_nth_row_value_by_prefix(rows, "(CH)", 1)
            )
            customer_data["contact_address_en"] = self._get_first_matching_value(
                rows,
                ["Contact Address (EN)", "Contact Address (EN) 客户公司英文地址"],
            )
            customer_data["contact_address_cn"] = self._get_nth_row_value_by_prefix(rows, "(CH)", 2)
            customer_data["delivery_address_cn"] = self._get_first_matching_value(
                rows,
                ["客户收件地址(CH)"],
            )

            postal_row = self._find_row(rows, "Postal code")
            if postal_row:
                postal_code = postal_row[1]["text"] if len(postal_row) > 1 else ""
                customer_data["postal_code"] = "" if postal_code in {"邮编", "Postal code"} else postal_code
                customer_data["telephone"] = self._get_label_value_pair(postal_row, ["*Telephone", "Telephone"])
                customer_data["fax"] = self._get_label_value_pair(postal_row, ["*Fax", "Fax"])
            else:
                customer_data["postal_code"] = ""
                customer_data["telephone"] = ""
                customer_data["fax"] = ""

            payment_row = self._find_row(rows, "Payment term")
            if payment_row:
                payment_options = payment_row[1]["checkbox_options"] if len(payment_row) > 1 else []
                customer_data["payment_term"] = self._first_checked_checkbox(payment_options)
                customer_data["payment_term_raw"] = payment_row[1]["text"] if len(payment_row) > 1 else ""
                if not customer_data["payment_term"]:
                    warnings.append("Payment term 未检测到勾选项，保留原始文本。")
            else:
                customer_data["payment_term"] = ""
                customer_data["payment_term_raw"] = ""

        if len(tables) > 3:
            contacts = self._extract_contacts(tables[3]["rows"])
            customer_data["contacts"] = contacts
            if contacts:
                first_contact = contacts[0]
                customer_data["contact_person"] = first_contact.get("name", "")
                customer_data["direct_line"] = first_contact.get("direct_line", "")
                customer_data["mobile"] = first_contact.get("mobile", "")
                customer_data["email"] = first_contact.get("email", "")
            else:
                customer_data["contact_person"] = ""
                customer_data["direct_line"] = ""
                customer_data["mobile"] = ""
                customer_data["email"] = ""
        else:
            customer_data["contacts"] = []

        if len(tables) > 4:
            table5_map = self._two_column_table_to_map(tables[4]["rows"])
            customer_data["hk_business_registration_no"] = table5_map.get(
                "Customers located in Hong Kong, pls. provide Business Registration NO.",
                "",
            )
            customer_data["tw_vat_no"] = table5_map.get(
                "Customers located in Taiwan, pls. provide VAT NO. (統一編號)",
                "",
            )
            customer_data["eu_vat_no"] = table5_map.get(
                "Customers located in Europe, pls. provide VAT NO.",
                "",
            )
            customer_data["other_tax_id"] = table5_map.get(
                "Indonesia，NPWP (Nomor Pokok Wajib Pajak，印尼纳税人识别号)",
                "",
            )

        if len(tables) > 5:
            rows = tables[5]["rows"]
            invoice_row = self._find_row_by_prefix(rows, "Invoice type")
            invoice_options = []
            if invoice_row:
                for cell in invoice_row[1:]:
                    invoice_options.extend(cell["checkbox_options"])
            customer_data["invoice_type"] = self._first_checked_checkbox(invoice_options)
            if not customer_data["invoice_type"]:
                warnings.append("Invoice type 未检测到勾选项。")

            table6_map = self._two_column_table_to_map(rows[1:])
            customer_data["vat_no"] = table6_map.get("VAT No.(纳税人识别号)", "")
            customer_data["bank_account"] = table6_map.get("Opening bank & Account No(银行全称/账号)", "")
            customer_data["registered_address_tel"] = table6_map.get(
                "Registered Address & tel. (注册地址/电话)",
                "",
            )

        if len(tables) > 6:
            table7_map = self._two_column_table_to_map(tables[6]["rows"])
            customer_data["requested_by"] = table7_map.get("Requested By", "")
            customer_data["requested_date"] = table7_map.get("Date", "")

        return {
            "customer_data": customer_data,
            "warnings": warnings,
        }

    def build_invoice_create_payload(self, customer_data: dict, config: dict) -> tuple[dict, list[str]]:
        """Build API payload for /soft-line/basic/invoice/create."""
        warnings = []
        bank_name, bank_account = self._split_bank_info(customer_data.get("bank_account", ""))
        invoice_type = self._map_invoice_type(customer_data)

        site_ids = self._map_site_ids(customer_data.get("branch_codes", []))
        if customer_data.get("branch_codes") and not site_ids:
            warnings.append(
                f"未配置 branch_codes={customer_data.get('branch_codes')} 对应的 siteIds。"
            )

        payer_name_cn = customer_data.get("customer_name_cn", "")
        payer_name_en = customer_data.get("customer_name_en", "")
        is_foreign_customer = not payer_name_cn.strip()
        contact_address_cn = customer_data.get("contact_address_cn", "")
        contact_address_en = customer_data.get("contact_address_en", "")
        delivery_address_cn = customer_data.get("delivery_address_cn", "")
        register_address = customer_data.get("registered_address_tel") or contact_address_cn or contact_address_en
        contact_address = delivery_address_cn or contact_address_cn or contact_address_en
        default_currency = "" if is_foreign_customer else "CNY"
        register_address_value = "" if is_foreign_customer else register_address
        register_tel_value = "" if is_foreign_customer else (
            customer_data.get("telephone")
            or customer_data.get("mobile")
            or customer_data.get("direct_line", "")
        )

        payload = OrderedDict(
            [
                ("customCode", "在建新客"),
                ("industryCode", customer_data.get("industry_code", "")),
                ("payerName", payer_name_cn or payer_name_en),
                ("payerNameEn", payer_name_en),
                ("taxPayerId", self._clean_tax_payer_id(customer_data.get("vat_no", ""))),
                (
                    "taxPayerIdEn",
                    customer_data.get("eu_vat_no")
                    or customer_data.get("tw_vat_no")
                    or customer_data.get("hk_business_registration_no")
                    or customer_data.get("other_tax_id", ""),
                ),
                ("bankName", bank_name),
                ("bankAccount", bank_account),
                ("registerAddress", register_address_value),
                ("registerTel", register_tel_value),
                ("registerFax", customer_data.get("fax", "")),
                ("invoiceType", invoice_type),
                ("monthlyPay", False),
                ("taxVat", ""),
                ("website", ""),
                ("shortName", payer_name_cn or payer_name_en),
                ("excludeRevenue", False),
                ("defaultCurrency", default_currency),
                ("isSystemSend", False),
                ("contactAddress", contact_address),
                ("siteIds", site_ids),
            ]
        )

        cleaned = {}
        for key, value in payload.items():
            keep = key in {"monthlyPay", "excludeRevenue", "isSystemSend"}
            if value is None:
                pass
            elif value == "":
                pass
            elif isinstance(value, list) and not value:
                pass
            else:
                keep = True
            if keep:
                cleaned[key] = value
        return cleaned, warnings

    def build_contact_payloads(self, customer_data: dict, invoice_id: int) -> list[dict]:
        """Build API payloads for /soft-line/basic/invoice/add/contacts."""
        address = (
            customer_data.get("delivery_address_cn")
            or customer_data.get("contact_address_cn")
            or customer_data.get("contact_address_en", "")
        )

        payloads = []
        for contact in customer_data.get("contacts", []):
            phone = contact.get("mobile") or contact.get("direct_line", "")
            if not any([contact.get("name"), phone, contact.get("email"), address]):
                continue
            payloads.append(
                {
                    "invoiceId": invoice_id,
                    "name": contact.get("name", ""),
                    "phone": phone,
                    "email": contact.get("email", ""),
                    "address": address,
                    "remark": "",
                }
            )
        return payloads

    def format_preview(
        self,
        parsed_form: dict,
        mapped: dict,
        invoice_payload: dict | None = None,
        contact_payloads: list | None = None,
    ) -> str:
        """Build a readable preview for the UI."""
        lines = [
            f"文件: {parsed_form['source_name']}",
            "",
            "ODM customer_data:",
            json.dumps(mapped["customer_data"], ensure_ascii=False, indent=2),
        ]
        if invoice_payload is not None:
            lines.extend(
                [
                    "",
                    "Invoice create payload:",
                    json.dumps(invoice_payload, ensure_ascii=False, indent=2),
                ]
            )
        if contact_payloads is not None:
            lines.extend(
                [
                    "",
                    "Contact payloads:",
                    json.dumps(contact_payloads, ensure_ascii=False, indent=2),
                ]
            )
        if mapped["warnings"]:
            lines.extend(["", "Warnings:", *[f"- {warning}" for warning in mapped["warnings"]]])
        return "\n".join(lines)

    def _extract_contacts(self, rows: list) -> list[dict]:
        headers = rows[0] if rows else []
        if len(headers) <= 1:
            return []

        contacts = []
        for col_idx in range(1, len(headers)):
            contact = {
                "name": self._get_contact_row_value(rows, "Name", col_idx),
                "direct_line": self._get_contact_row_value(rows, "Direct Line", col_idx),
                "mobile": self._get_contact_row_value(rows, "Mobile", col_idx),
                "email": self._get_contact_row_value(rows, "Email", col_idx),
            }
            if any(contact.values()):
                contacts.append(contact)
        return contacts

    def _get_contact_row_value(self, rows: list, label: str, col_idx: int) -> str:
        row = self._find_row(rows, label)
        if not row or len(row) <= col_idx:
            return ""
        return row[col_idx]["text"]

    def _map_invoice_type(self, customer_data: dict):
        label = customer_data.get("invoice_type", "").strip()
        vat_no = customer_data.get("vat_no", "")
        is_foreign_customer = not customer_data.get("customer_name_cn", "").strip()

        if "不需要发票" in vat_no or "无需发票" in vat_no:
            return self.INVOICE_TYPE_NO_NEED
        if label == "专票":
            return self.INVOICE_TYPE_SPECIAL
        if label == "普票":
            return self.INVOICE_TYPE_NORMAL
        if is_foreign_customer:
            return self.INVOICE_TYPE_INVOICE
        return self.INVOICE_TYPE_INVOICE

    def _map_site_ids(self, branch_codes: list[str]) -> list[int]:
        site_ids = []
        for branch_code in branch_codes:
            site_id = self.SITE_ID_MAP.get(branch_code.strip())
            if site_id is not None and site_id not in site_ids:
                site_ids.append(site_id)
        return site_ids

    def _clean_tax_payer_id(self, value: str) -> str:
        text = (value or "").strip()
        if not text:
            return ""
        text = re.sub(r"[（(]\s*(不需要发票|无需发票)\s*[）)]", "", text)
        return text.strip()

    def _clean_company_name(self, value: str) -> str:
        text = (value or "").strip()
        return re.sub(r"(?<!\d)[123]$", "", text).strip()

    def _split_bank_info(self, bank_info: str) -> tuple[str, str]:
        text = (bank_info or "").strip()
        if not text:
            return "", ""

        match = re.match(r"^(.*?)([0-9A-Za-z]{8,})$", text)
        if not match:
            return text, ""
        return match.group(1).strip(), match.group(2).strip()

    def _get_first_matching_value(self, rows: list, prefixes: list[str]) -> str:
        for prefix in prefixes:
            row = self._find_row_by_prefix(rows, prefix)
            if row and len(row) > 1:
                return row[1]["text"]
        return ""

    def _get_nth_row_value_by_prefix(self, rows: list, prefix: str, occurrence: int) -> str:
        count = 0
        for row in rows:
            if row and row[0]["text"].startswith(prefix):
                count += 1
                if count == occurrence:
                    return row[1]["text"] if len(row) > 1 else ""
        return ""

    def _get_label_value_pair(self, row: list, labels: list[str]) -> str:
        for idx, cell in enumerate(row[:-1]):
            if cell["text"] in labels:
                return row[idx + 1]["text"]
        return ""

    def _two_column_table_to_map(self, rows: list) -> dict:
        values = {}
        for row in rows:
            if len(row) >= 2 and row[0]["text"]:
                values[row[0]["text"]] = row[1]["text"]
        return values

    def _find_row(self, rows: list, label: str):
        for row in rows:
            if row and row[0]["text"] == label:
                return row
        return None

    def _find_row_by_prefix(self, rows: list, prefix: str):
        for row in rows:
            if row and row[0]["text"].startswith(prefix):
                return row
        return None

    def _get_row_value(self, rows: list, label: str, index: int) -> str:
        row = self._find_row(rows, label)
        if not row or len(row) <= index:
            return ""
        return row[index]["text"]

    def _first_checked_label(self, table: dict) -> str:
        labels = self._all_checked_labels(table)
        return labels[0] if labels else ""

    def _all_checked_labels(self, table: dict) -> list[str]:
        labels = []
        for row in table["rows"]:
            for cell in row:
                for option in cell["checkbox_options"]:
                    if option["checked"] and option["label"] not in labels:
                        labels.append(option["label"])
        return labels

    def _first_checked_checkbox(self, options: list) -> str:
        for option in options:
            if option["checked"]:
                return option["label"]
        return ""
