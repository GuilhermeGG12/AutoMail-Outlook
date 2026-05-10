from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from mailmerge_assistant.models import ReportRow
from mailmerge_assistant.report_writer import REPORT_HEADERS, write_report


def test_write_report_creates_xlsx_with_rows(tmp_path: Path) -> None:
    path = write_report(
        [
            ReportRow(
                linha=2,
                razao_social="Empresa",
                proprietario="Maria",
                emails="cliente@email.com",
                assunto="Honorários contábeis - Empresa",
                valor="R$ 300,00",
                dia_pagamento="10",
                anexo="arquivo.pdf",
                status="OK",
                mensagem="OK",
            )
        ],
        reports_dir=tmp_path,
    )

    assert path.exists()
    workbook = load_workbook(path)
    worksheet = workbook.active
    assert [cell.value for cell in worksheet[1]] == REPORT_HEADERS
    assert worksheet["B2"].value == 2
    assert worksheet["J2"].value == "OK"
