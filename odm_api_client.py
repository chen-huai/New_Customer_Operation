# -*- coding: utf-8 -*-
"""ODM API client."""

from __future__ import annotations

from typing import Iterable

import requests


class OdmApiClient:
    """Call ODM APIs with config-driven environment selection."""

    def __init__(self, config: dict):
        self.config = config
        self.base_url = self._resolve_base_url(config)
        self.session = requests.Session()
        self.token = None

    def login(self) -> str:
        """Login and cache the access token."""
        username = self.config.get("Account", "").strip()
        password = self.config.get("Password", "").strip()
        response = self.session.post(
            self._url("/soft-line/auth/login"),
            json={
                "username": username,
                "password": password,
            },
            timeout=30,
        )
        try:
            payload = self._parse_response(response)
        except Exception as exc:
            raise RuntimeError(
                f"{exc}；environment={self.config.get('Environment', 'test').strip().lower()}；"
                f"base_url={self.base_url}；username={username}"
            ) from exc
        token = (payload.get("data") or {}).get("accessToken")
        if not token:
            raise RuntimeError("登录成功但未返回 accessToken")

        self.token = token
        self.session.headers.update(
            {
                "Authorization": f"Bearer {token}",
                "accessToken": token,
            }
        )
        return token

    def create_invoice(self, invoice_payload: dict) -> dict:
        """Create the customer/invoice record."""
        self._ensure_login()
        response = self.session.post(
            self._url("/soft-line/basic/invoice/create"),
            json=invoice_payload,
            timeout=30,
        )
        payload = self._parse_response(response)
        data = payload.get("data") or {}
        invoice_id = data.get("id")
        if invoice_id is None:
            raise RuntimeError("客户创建成功，但返回结果中没有 id")
        return data

    def add_contacts(self, contacts_payload: Iterable[dict]) -> dict:
        """Create contacts for an invoice."""
        self._ensure_login()
        response = self.session.post(
            self._url("/soft-line/basic/invoice/add/contacts"),
            json=list(contacts_payload),
            timeout=30,
        )
        return self._parse_response(response)

    def _ensure_login(self) -> None:
        if not self.token:
            self.login()

    def _url(self, path: str) -> str:
        return f"{self.base_url}{path}"

    def _resolve_base_url(self, config: dict) -> str:
        environment = config.get("Environment", "test").strip().lower()
        if environment == "prod":
            base_url = config.get("Prod_API_Base_URL", "").strip()
        else:
            base_url = config.get("Test_API_Base_URL", "").strip()

        if not base_url:
            raise RuntimeError(f"未配置 {environment} 环境的 API 根地址")
        return base_url.rstrip("/")

    def _parse_response(self, response: requests.Response) -> dict:
        try:
            payload = response.json()
        except ValueError as exc:
            raise RuntimeError(f"接口返回的不是 JSON: HTTP {response.status_code}") from exc

        if response.status_code >= 400:
            message = payload.get("message") or response.text
            raise RuntimeError(f"HTTP {response.status_code}: {message}")

        if payload.get("success") is False:
            raise RuntimeError(payload.get("message") or "接口返回 success=false")

        return payload
