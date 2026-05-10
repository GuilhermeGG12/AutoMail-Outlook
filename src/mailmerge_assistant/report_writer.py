from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

from mailmerge_assistant.config import REPORTS_DIR
from mailmerge_assistant.models import ReportRow

REPORT_HEADERS = [
    "Data_Hora",
    "Linha",
    "Razão Social",
    "Proprietário/Dirigente",
    "Emails",
    "Assunto",
    "Valor",
    "Dia de Pagamento",
    "Anexo",
    "Status",
    "Mensagem",
]


def write_report(rows: list[ReportRow], reports_dir: Path = REPORTS_DIR) -> Path:
    reports_dir.mkdir(parents=True, exist_ok=True)
    now = datetime.now()
    path = reports_dir / f"relatorio_envio_{now:%Y%m%d_%H%M%S}.xlsx"

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Relatório"
    worksheet.append(REPORT_HEADERS)
    data_hora = now.strftime("%Y-%m-%d %H:%M:%S")
    for row in rows:
        worksheet.append(
            [
                data_hora,
                row.linha,
                row.razao_social,
                row.proprietario,
                row.emails,
                row.assunto,
                row.valor,
                row.dia_pagamento,
                row.anexo,
                row.status,
                row.mensagem,
            ]
        )
    for column in worksheet.columns:
        max_length = max(len(str(cell.value or "")) for cell in column)
        worksheet.column_dimensions[column[0].column_letter].width = min(max_length + 2, 60)
    workbook.save(path)
    return path
