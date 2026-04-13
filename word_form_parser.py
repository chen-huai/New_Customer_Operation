# -*- coding: utf-8 -*-
"""Parse customer Word forms in .docx and .doc format."""

from __future__ import annotations

import os
import subprocess
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile


WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


class WordFormParser:
    """Parse structured Word forms into tables and checkbox data."""

    def parse(self, file_path: str) -> dict:
        """Parse a .docx or .doc file."""
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"文件不存在: {file_path}")

        suffix = path.suffix.lower()
        if suffix == ".docx":
            return self._parse_docx(path)
        if suffix == ".doc":
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_docx = Path(temp_dir) / f"{path.stem}.docx"
                self._convert_doc_to_docx(path, temp_docx)
                return self._parse_docx(temp_docx, source_file=path)
        raise ValueError("仅支持 .docx 或 .doc 文件")

    def _parse_docx(self, path: Path, source_file: Path | None = None) -> dict:
        with ZipFile(path) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))

        tables = []
        for table_index, table in enumerate(root.findall(".//w:tbl", WORD_NS), start=1):
            rows = []
            for row in table.findall("./w:tr", WORD_NS):
                cells = [self._parse_cell(cell) for cell in row.findall("./w:tc", WORD_NS)]
                if any(cell["text"] for cell in cells):
                    rows.append(cells)
            if rows:
                tables.append({"table_index": table_index, "rows": rows})

        return {
            "source_file": str(source_file or path),
            "source_name": (source_file or path).name,
            "tables": tables,
        }

    def _parse_cell(self, cell) -> dict:
        text = self._paragraph_text(cell)
        checkbox_options = []

        for paragraph in cell.findall("./w:p", WORD_NS):
            checkboxes = paragraph.findall(".//w:checkBox", WORD_NS)
            if not checkboxes:
                continue

            label = self._paragraph_text(paragraph).strip()
            if not label:
                continue

            checked = any(self._is_checkbox_checked(checkbox) for checkbox in checkboxes)
            checkbox_options.append({"label": label, "checked": checked})

        return {
            "text": text.strip(),
            "checkbox_options": checkbox_options,
        }

    def _paragraph_text(self, element) -> str:
        parts = []
        for text_node in element.findall(".//w:t", WORD_NS):
            parts.append(text_node.text or "")
        return "".join(parts)

    def _is_checkbox_checked(self, checkbox) -> bool:
        checked_node = checkbox.find("./w:checked", WORD_NS)
        if checked_node is None:
            return False

        value = checked_node.get(f"{{{WORD_NS['w']}}}val")
        return value not in {"0", "false", "False"}

    def _convert_doc_to_docx(self, input_path: Path, output_path: Path) -> None:
        input_value = self._ps_quote(str(input_path.resolve()))
        output_value = self._ps_quote(str(output_path.resolve()))
        script = f"""
$word = $null
$document = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $document = $word.Documents.Open({input_value})
    $document.SaveAs([ref] {output_value}, [ref] 16)
}} finally {{
    if ($document -ne $null) {{ $document.Close() }}
    if ($word -ne $null) {{ $word.Quit() }}
}}
"""
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", script],
            check=False,
            capture_output=True,
            text=True,
        )
        if result.returncode != 0 or not output_path.exists():
            stderr = (result.stderr or result.stdout or "").strip()
            raise RuntimeError(f".doc 转换失败: {stderr}")

    @staticmethod
    def _ps_quote(value: str) -> str:
        return "'" + value.replace("'", "''") + "'"
