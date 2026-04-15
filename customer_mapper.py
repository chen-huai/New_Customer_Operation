# -*- coding: utf-8 -*-
"""Map parsed Word form data into normalized customer fields and ODM payloads."""

from __future__ import annotations

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
    NO_INVOICE_KEYWORDS = (
        "不需要发票",
        "无需发票",
        "无须发票",
        "不要发票",
        "不要票",
        "不用发票",
        "不用票",
        "不开发票",
        "不开票",
        "无需开票",
        "无须开票",
        "不要开票",
        "免开发票",
        "免开票",
        "免票",
        "无需票",
        "无须票",
        "不用开票",
        "不要税票",
        "不用税票",
        "无需税票",
        "无须税票",
    )
    NO_INVOICE_PATTERN = re.compile(
        r"[\s,，;；]*[（(【\[]?\s*(?:"
        + "|".join(re.escape(keyword) for keyword in NO_INVOICE_KEYWORDS)
        + r")\s*[）)】\]]?[\s,，;；]*"
    )
    LABEL_NORMALIZE_PATTERN = re.compile(r"[\s\-_/\\:：,.，;；&()（）【】\[\]<>]+")
    CONTACT_FIELD_ALIASES = {
        "name": ["name", "姓名", "联系人姓名"],
        "direct_line": ["directline", "电话", "直线", "联系电话"],
        "mobile": ["mobile", "手机", "手机号"],
        "email": ["email", "邮箱", "电子邮箱"],
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
            warnings.append("未检测到 Customer category 勾选项。")

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
                    [
                        "Customer Name (EN)",
                        "Customer Name (EN)客户公司英文名称",
                        "customernameen",
                        "客户名称英文",
                        "客户公司英文名称",
                    ],
                )
            )
            customer_data["customer_name_cn"] = self._clean_company_name(
                self._get_customer_name_cn(rows)
            )
            customer_data["contact_address_en"] = self._get_first_matching_value(
                rows,
                [
                    "Contact Address (EN)",
                    "Contact Address (EN) 客户公司英文地址",
                    "contactaddressen",
                    "联系地址英文",
                    "客户公司英文地址",
                ],
            )
            customer_data["contact_address_cn"] = self._get_contact_address_cn(rows)
            customer_data["delivery_address_cn"] = self._get_first_matching_value(
                rows,
                [
                    "客户收件地址(CH)",
                    "Delivery Address (CH)",
                    "deliveryaddressch",
                    "deliveryaddresscn",
                    "客户收件地址",
                    "收件地址中文",
                ],
            )

            postal_row = self._find_row_by_aliases(rows, ["Postal code", "postalcode", "邮政编码", "邮编"])
            if postal_row:
                postal_code = postal_row[1]["text"] if len(postal_row) > 1 else ""
                customer_data["postal_code"] = "" if self._normalize_label(postal_code) in {"邮编", "postalcode"} else postal_code
                customer_data["telephone"] = self._get_label_value_pair(
                    postal_row,
                    ["*Telephone", "Telephone", "telephone", "电话", "联系电话"],
                )
                customer_data["fax"] = self._get_label_value_pair(
                    postal_row,
                    ["*Fax", "Fax", "fax", "传真"],
                )
            else:
                customer_data["postal_code"] = ""
                customer_data["telephone"] = ""
                customer_data["fax"] = ""

            payment_row = self._find_row_by_aliases(rows, ["Payment term", "paymentterm", "付款条件"])
            if payment_row:
                payment_options = []
                for cell in payment_row[1:]:
                    payment_options.extend(cell["checkbox_options"])
                customer_data["payment_term"] = self._first_checked_checkbox(payment_options)
                customer_data["payment_term_raw"] = self._join_row_values(payment_row[1:])
                if not customer_data["payment_term"]:
                    warnings.append("未检测到 Payment term 勾选项，已保留原始文本。")
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
            rows = tables[4]["rows"]
            customer_data["hk_business_registration_no"] = self._get_two_column_value(
                rows,
                [
                    "Customers located in Hong Kong, pls. provide Business Registration NO.",
                    "hongkongbusinessregistrationno",
                    "businessregistrationno",
                ],
            )
            customer_data["tw_vat_no"] = self._get_two_column_value(
                rows,
                [
                    "Customers located in Taiwan, pls. provide VAT NO. (統一編號)",
                    "taiwanvatno",
                    "統一編號",
                ],
            )
            customer_data["eu_vat_no"] = self._get_two_column_value(
                rows,
                [
                    "Customers located in Europe, pls. provide VAT NO.",
                    "europevatno",
                    "customerslocatedineurope",
                ],
            )
            customer_data["other_tax_id"] = self._get_two_column_value(
                rows,
                [
                    "Indonesia，NPWP (Nomor Pokok Wajib Pajak，印尼纳税人识别号)",
                    "indonesianpwp",
                    "npwp",
                    "印尼纳税人识别号",
                ],
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
                warnings.append("未检测到 Invoice type 勾选项。")

            customer_data["vat_no"] = self._get_two_column_value(
                rows[1:],
                ["VAT No.(纳税人识别号)", "vatno", "纳税人识别号"],
            )
            customer_data["bank_account"] = self._get_two_column_value(
                rows[1:],
                [
                    "Opening bank & Account No(银行全称/账号)",
                    "openingbankaccountno",
                    "银行全称账号",
                    "开户行账号",
                ],
            )
            customer_data["registered_address_tel"] = self._get_two_column_value(
                rows[1:],
                [
                    "Registered Address & tel. (注册地址/电话)",
                    "registeredaddresstel",
                    "注册地址电话",
                ],
            )

        if len(tables) > 6:
            rows = tables[6]["rows"]
            customer_data["requested_by"] = self._get_two_column_value(
                rows,
                ["Requested By", "requestedby", "申请人", "提交人"],
            )
            customer_data["requested_date"] = self._get_two_column_value(
                rows,
                ["Date", "date", "日期"],
            )

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
        register_address, register_tel = self._split_registered_address_tel(
            customer_data.get("registered_address_tel", "")
        )
        register_address = register_address or contact_address_cn or contact_address_en
        contact_address = delivery_address_cn or contact_address_cn or contact_address_en
        default_currency = "" if is_foreign_customer else "CNY"
        domestic_tax_payer_id = self._clean_tax_payer_id(customer_data.get("vat_no", ""))
        foreign_tax_payer_id = self._get_foreign_tax_payer_id(customer_data)
        bank_name_value = "" if is_foreign_customer else bank_name
        bank_account_value = "" if is_foreign_customer else bank_account
        register_address_value = "" if is_foreign_customer else register_address
        register_tel_value = "" if is_foreign_customer else (
            register_tel
            or customer_data.get("telephone")
            or customer_data.get("mobile")
            or customer_data.get("direct_line", "")
        )
        register_fax_value = "" if is_foreign_customer else customer_data.get("fax", "")
        website_value = customer_data.get("website", "") if is_foreign_customer else ""

        payload = OrderedDict(
            [
                ("customCode", "在建新客"),
                ("industryCode", customer_data.get("industry_code", "")),
                ("payerName", payer_name_cn or payer_name_en),
                ("payerNameEn", payer_name_en),
                ("taxPayerId", "" if is_foreign_customer else domestic_tax_payer_id),
                ("taxPayerIdEn", foreign_tax_payer_id if is_foreign_customer else ""),
                ("bankName", bank_name_value),
                ("bankAccount", bank_account_value),
                ("registerAddress", register_address_value),
                ("registerTel", register_tel_value),
                ("registerFax", register_fax_value),
                ("invoiceType", invoice_type),
                ("monthlyPay", False),
                ("taxVat", ""),
                ("website", website_value),
                ("shortName", payer_name_cn or payer_name_en),
                ("excludeRevenue", False),
                ("defaultCurrency", default_currency),
                ("isSystemSend", True),
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

    def format_payer_preview(
        self,
        parsed_form: dict,
        mapped: dict,
        invoice_payload: dict,
        contact_payloads: list[dict],
        extra_warnings: list[str] | None = None,
    ) -> str:
        """Build a payer-only preview for the UI."""
        lines = self._build_preview_lines(parsed_form, mapped)
        lines.extend(
            [
                "",
                "付款方预览:",
                f"ODM名称: {self._display_value(invoice_payload.get('payerName', ''))}",
                f"英文名称: {self._display_value(invoice_payload.get('payerNameEn', ''))}",
                f"国内税号: {self._display_value(invoice_payload.get('taxPayerId', ''))}",
                f"国外税号: {self._display_value(invoice_payload.get('taxPayerIdEn', ''))}",
                f"站点ID: {self._display_value(', '.join(str(item) for item in invoice_payload.get('siteIds', [])))}",
                f"默认币种: {self._display_value(invoice_payload.get('defaultCurrency', ''))}",
                f"待建联系人: {len(contact_payloads)}",
            ]
        )
        self._append_warning_lines(lines, mapped, extra_warnings)
        return "\n".join(lines)

    def format_applicant_preview(
        self,
        parsed_form: dict,
        mapped: dict,
        applicant_payload: dict,
        extra_warnings: list[str] | None = None,
    ) -> str:
        """Build an applicant-only preview for the UI."""
        lines = self._build_preview_lines(parsed_form, mapped)
        lines.extend(
            [
                "",
                "申请方预览:",
                f"申请方名称: {self._display_value(applicant_payload.get('name', ''))}",
                f"申请方英文名: {self._display_value(applicant_payload.get('nameEn', ''))}",
            ]
        )
        self._append_warning_lines(lines, mapped, extra_warnings)
        return "\n".join(lines)

    def build_applicant_create_payload(self, customer_data: dict) -> tuple[dict, list[str]]:
        """Build API payload for /soft-line/basic/applicant/create."""
        warnings = []
        payer_name_cn = customer_data.get("customer_name_cn", "").strip()
        payer_name_en = customer_data.get("customer_name_en", "").strip()
        applicant_name = payer_name_cn or payer_name_en

        if not applicant_name:
            warnings.append("未获取到付款中英文名称，无法创建 applicant。")
            return {}, warnings

        payload = OrderedDict(
            [
                ("name", applicant_name),
                ("nameEn", payer_name_en),
                ("remark", ""),
            ]
        )

        cleaned = {}
        for key, value in payload.items():
            if value is None or value == "":
                continue
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
        applicant_payload: dict | None = None,
        invoice_payload: dict | None = None,
        contact_payloads: list | None = None,
        extra_warnings: list[str] | None = None,
    ) -> str:
        """Build a concise preview for the UI."""
        customer_data = mapped["customer_data"]
        lines = [
            f"文件: {parsed_form['source_name']}",
            f"分公司: {self._display_value(' / '.join(customer_data.get('branch_codes', [])))}",
            f"客户类别: {self._display_value(customer_data.get('customer_category', ''))}",
            f"客户名称: {self._format_bilingual_name(customer_data)}",
            f"订单金额: {self._format_amount(customer_data)}",
            f"付款条件: {self._display_value(customer_data.get('payment_term') or customer_data.get('payment_term_raw', ''))}",
            f"发票类型: {self._display_value(customer_data.get('invoice_type', ''))}",
            f"联系地址: {self._display_value(self._primary_address(customer_data))}",
            f"主要联系人: {self._format_primary_contact(customer_data)}",
            f"联系人数量: {len(customer_data.get('contacts', []))}",
        ]

        if applicant_payload is not None:
            lines.extend(
                [
                    "",
                    "提交摘要:",
                    f"Applicant: {self._display_value(applicant_payload.get('name', ''))}",
                ]
            )
        if invoice_payload is not None:
            lines.extend(
                [
                    f"ODM名称: {self._display_value(invoice_payload.get('payerName', ''))}",
                    f"英文名称: {self._display_value(invoice_payload.get('payerNameEn', ''))}",
                    f"站点ID: {self._display_value(', '.join(str(item) for item in invoice_payload.get('siteIds', [])))}",
                    f"默认币种: {self._display_value(invoice_payload.get('defaultCurrency', ''))}",
                ]
            )
        if contact_payloads is not None:
            lines.append(f"待建联系人: {len(contact_payloads)}")

        warnings = list(mapped["warnings"])
        if extra_warnings:
            warnings.extend(extra_warnings)
        if warnings:
            lines.extend(["", "提醒:"])
            lines.extend(f"- {warning}" for warning in warnings)
        return "\n".join(lines)

    def format_submit_result(
        self,
        customer_data: dict,
        environment: str,
        invoice_id: int,
        created_contacts: int,
        warnings: list[str] | None = None,
    ) -> str:
        lines = [
            "提交结果:",
            f"环境: {environment}",
            f"客户名称: {self._format_bilingual_name(customer_data)}",
            f"invoiceId: {invoice_id}",
            f"联系人数量: {created_contacts}",
        ]
        if warnings:
            lines.extend(["", "提醒:"])
            lines.extend(f"- {warning}" for warning in warnings)
        return "\n".join(lines)

    def format_payer_submit_result(
        self,
        customer_data: dict,
        environment: str,
        invoice_id: int | None,
        created_contacts: int,
        warnings: list[str] | None = None,
        status_text: str = "创建成功",
    ) -> str:
        """Build a payer-only submit result for the UI."""
        lines = [
            "付款方提交结果:",
            f"状态: {status_text}",
            f"环境: {environment}",
            f"客户名称: {self._format_bilingual_name(customer_data)}",
        ]
        if invoice_id is not None:
            lines.append(f"invoiceId: {invoice_id}")
        lines.append(f"联系人数量: {created_contacts}")
        if warnings:
            lines.extend(["", "提醒:"])
            lines.extend(f"- {warning}" for warning in warnings)
        return "\n".join(lines)

    def format_applicant_submit_result(
        self,
        customer_data: dict,
        environment: str,
        applicant_payload: dict,
        warnings: list[str] | None = None,
        status_text: str = "创建成功",
    ) -> str:
        """Build an applicant-only submit result for the UI."""
        lines = [
            "申请方提交结果:",
            f"状态: {status_text}",
            f"环境: {environment}",
            f"客户名称: {self._format_bilingual_name(customer_data)}",
            f"申请方名称: {self._display_value(applicant_payload.get('name', ''))}",
            f"申请方英文名: {self._display_value(applicant_payload.get('nameEn', ''))}",
        ]
        if warnings:
            lines.extend(["", "提醒:"])
            lines.extend(f"- {warning}" for warning in warnings)
        return "\n".join(lines)

    def _extract_contacts(self, rows: list) -> list[dict]:
        headers = rows[0] if rows else []
        if len(headers) <= 1:
            return []

        contacts = []
        for col_idx in range(1, len(headers)):
            contact = {
                "name": self._get_contact_row_value(rows, "name", col_idx),
                "direct_line": self._get_contact_row_value(rows, "direct_line", col_idx),
                "mobile": self._get_contact_row_value(rows, "mobile", col_idx),
                "email": self._get_contact_row_value(rows, "email", col_idx),
            }
            if any(contact.values()):
                contacts.append(contact)
        return contacts

    def _get_contact_row_value(self, rows: list, field_name: str, col_idx: int) -> str:
        row = self._find_row_by_aliases(rows, self.CONTACT_FIELD_ALIASES.get(field_name, [field_name]))
        if not row or len(row) <= col_idx:
            return ""
        return row[col_idx]["text"]

    def _map_invoice_type(self, customer_data: dict):
        label = customer_data.get("invoice_type", "").strip()
        vat_no = customer_data.get("vat_no", "")
        is_foreign_customer = not customer_data.get("customer_name_cn", "").strip()

        if self._contains_no_invoice_flag(vat_no):
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
        text = self.NO_INVOICE_PATTERN.sub("", text)
        return text.strip()

    def _contains_no_invoice_flag(self, value: str) -> bool:
        text = (value or "").strip()
        return any(keyword in text for keyword in self.NO_INVOICE_KEYWORDS)

    def _get_foreign_tax_payer_id(self, customer_data: dict) -> str:
        return (
            customer_data.get("eu_vat_no")
            or customer_data.get("tw_vat_no")
            or customer_data.get("hk_business_registration_no")
            or customer_data.get("other_tax_id", "")
        ).strip()

    def _clean_company_name(self, value: str) -> str:
        text = (value or "").strip()
        return re.sub(r"(?<!\d)[123]$", "", text).strip()

    def _display_value(self, value: str) -> str:
        text = (value or "").strip()
        return text or "-"

    def _build_preview_lines(self, parsed_form: dict, mapped: dict) -> list[str]:
        customer_data = mapped["customer_data"]
        return [
            f"文件: {parsed_form['source_name']}",
            f"分公司: {self._display_value(' / '.join(customer_data.get('branch_codes', [])))}",
            f"客户类别: {self._display_value(customer_data.get('customer_category', ''))}",
            f"客户名称: {self._format_bilingual_name(customer_data)}",
            f"订单金额: {self._format_amount(customer_data)}",
            f"付款条件: {self._display_value(customer_data.get('payment_term') or customer_data.get('payment_term_raw', ''))}",
            f"发票类型: {self._display_value(customer_data.get('invoice_type', ''))}",
            f"联系地址: {self._display_value(self._primary_address(customer_data))}",
            f"主要联系人: {self._format_primary_contact(customer_data)}",
            f"联系人数量: {len(customer_data.get('contacts', []))}",
        ]

    def _append_warning_lines(
        self,
        lines: list[str],
        mapped: dict,
        extra_warnings: list[str] | None = None,
    ) -> None:
        warnings = list(mapped["warnings"])
        if extra_warnings:
            warnings.extend(extra_warnings)
        if warnings:
            lines.extend(["", "提醒:"])
            lines.extend(f"- {warning}" for warning in warnings)

    def _format_bilingual_name(self, customer_data: dict) -> str:
        payer_name_cn = (customer_data.get("customer_name_cn") or "").strip()
        payer_name_en = (customer_data.get("customer_name_en") or "").strip()
        if payer_name_cn and payer_name_en:
            return f"{payer_name_cn} / {payer_name_en}"
        return self._display_value(payer_name_cn or payer_name_en)

    def _format_amount(self, customer_data: dict) -> str:
        amount = (customer_data.get("new_order_value") or "").strip()
        currency = (customer_data.get("currency") or "").strip()
        if amount and currency:
            return f"{amount} {currency}"
        return self._display_value(amount or currency)

    def _primary_address(self, customer_data: dict) -> str:
        return (
            customer_data.get("delivery_address_cn")
            or customer_data.get("contact_address_cn")
            or customer_data.get("contact_address_en")
            or ""
        )

    def _format_primary_contact(self, customer_data: dict) -> str:
        contact = (customer_data.get("contact_person") or "").strip()
        phone = (
            customer_data.get("mobile")
            or customer_data.get("direct_line")
            or customer_data.get("telephone")
            or ""
        ).strip()
        email = (customer_data.get("email") or "").strip()

        parts = [item for item in [contact, phone, email] if item]
        if not parts:
            return "-"
        return " / ".join(parts)

    def _split_bank_info(self, bank_info: str) -> tuple[str, str]:
        text = (bank_info or "").strip()
        if not text:
            return "", ""
        normalized = re.sub(r"[\r\n]+", "\n", text)
        parts = [
            part.strip(" \t\r\n/|;；,，&")
            for part in re.split(r"\n+|/|\\|\||;|；|,|，|&", normalized)
            if part.strip(" \t\r\n/|;；,，&")
        ]

        account_index = None
        account_value = ""
        for idx in range(len(parts) - 1, -1, -1):
            candidate = re.sub(r"\s+", "", parts[idx])
            if re.fullmatch(r"[0-9A-Za-z\-]{8,}", candidate):
                account_index = idx
                account_value = candidate
                break

        if account_index is not None:
            bank_name_parts = [part for i, part in enumerate(parts) if i != account_index]
            return " ".join(bank_name_parts).strip(), account_value

        compact = re.sub(r"\s+", " ", normalized.replace("\n", " ")).strip()
        matches = list(re.finditer(r"[0-9A-Za-z\-]{8,}", compact))
        if not matches:
            return compact, ""

        account_match = matches[-1]
        bank_name = (
            compact[: account_match.start()] + " " + compact[account_match.end() :]
        ).strip(" /|,，;；\\&")
        bank_name = re.sub(r"\s+", " ", bank_name).strip()
        return bank_name, account_match.group(0).strip()

    def _split_registered_address_tel(self, value: str) -> tuple[str, str]:
        text = (value or "").strip()
        if not text:
            return "", ""

        normalized = re.sub(r"[\r\n]+", "\n", text)
        phone_pattern = re.compile(r"(?<!\d)(?:\+?\d[\d\s\-]{6,}\d)(?!\d)")

        parts = [
            part.strip(" \t\r\n/|;；,，&")
            for part in re.split(r"\n+|/|\\|\||;|；|,|，|&", normalized)
            if part.strip(" \t\r\n/|;；,，&")
        ]

        phone_index = None
        phone_value = ""
        for idx in range(len(parts) - 1, -1, -1):
            match = phone_pattern.search(parts[idx])
            if match:
                phone_index = idx
                phone_value = re.sub(r"\s+", "", match.group(0))
                break

        if phone_index is not None:
            address_parts = [part for i, part in enumerate(parts) if i != phone_index]
            remainder = phone_pattern.sub("", parts[phone_index]).strip(" /|,，;；\\&")
            if remainder:
                address_parts.insert(phone_index, remainder)
            return " ".join(address_parts).strip(), phone_value

        compact = re.sub(r"\s+", " ", normalized.replace("\n", " ")).strip()
        matches = list(phone_pattern.finditer(compact))
        if not matches:
            return compact, ""

        last_match = matches[-1]
        phone_value = re.sub(r"\s+", "", last_match.group(0))
        address = (
            compact[: last_match.start()] + " " + compact[last_match.end() :]
        ).strip(" /|,，;；\\&")
        return re.sub(r"\s+", " ", address).strip(), phone_value

    def _get_first_matching_value(self, rows: list, prefixes: list[str]) -> str:
        row = self._find_row_by_aliases(rows, prefixes)
        if row and len(row) > 1:
            return row[1]["text"]
        return ""

    def _get_customer_name_cn(self, rows: list) -> str:
        value = self._get_first_matching_value(
            rows,
            [
                "Customer Name (CH)",
                "Customer Name(CH)",
                "(CH) 客户公司中文名称",
                "(CH)客户公司中文名称",
                "customernamech",
                "customernamecn",
                "客户名称中文",
                "客户公司中文名称",
            ],
        )
        if value:
            return value
        return self._get_nth_row_value_by_aliases(rows, ["(CH)", "中文"], 1)

    def _get_contact_address_cn(self, rows: list) -> str:
        value = self._get_first_matching_value(
            rows,
            [
                "Contact Address (CH)",
                "Contact Address(CH)",
                "(CH) 客户公司中文地址",
                "(CH)客户公司中文地址",
                "contactaddressch",
                "contactaddresscn",
                "联系地址中文",
                "客户公司中文地址",
            ],
        )
        if value:
            return value
        return self._get_nth_row_value_by_aliases(rows, ["(CH)", "中文"], 2)

    def _get_nth_row_value_by_aliases(self, rows: list, aliases: list[str], occurrence: int) -> str:
        count = 0
        for row in rows:
            if row and self._row_matches_aliases(row, aliases):
                count += 1
                if count == occurrence:
                    return row[1]["text"] if len(row) > 1 else ""
        return ""

    def _get_label_value_pair(self, row: list, labels: list[str]) -> str:
        for idx, cell in enumerate(row[:-1]):
            if self._text_matches_aliases(cell["text"], labels):
                return row[idx + 1]["text"]
        return ""

    def _get_two_column_value(self, rows: list, aliases: list[str]) -> str:
        row = self._find_row_by_aliases(rows, aliases)
        if row and len(row) > 1:
            return row[1]["text"]
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

    def _find_row_by_aliases(self, rows: list, aliases: list[str]):
        for row in rows:
            if row and self._row_matches_aliases(row, aliases):
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

    def _join_row_values(self, cells: list[dict]) -> str:
        parts = [cell.get("text", "").strip() for cell in cells if cell.get("text", "").strip()]
        return " ".join(parts)

    def _row_matches_aliases(self, row: list, aliases: list[str]) -> bool:
        return bool(row) and self._text_matches_aliases(row[0]["text"], aliases)

    def _text_matches_aliases(self, text: str, aliases: list[str]) -> bool:
        normalized_text = self._normalize_label(text)
        if not normalized_text:
            return False
        for alias in aliases:
            normalized_alias = self._normalize_label(alias)
            if normalized_alias and normalized_alias in normalized_text:
                return True
        return False

    def _normalize_label(self, value: str) -> str:
        text = (value or "").strip().lower()
        return self.LABEL_NORMALIZE_PATTERN.sub("", text)
