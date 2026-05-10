from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from mailmerge_assistant.excel_reader import (
    ExcelReaderError,
    missing_required_columns,
    read_clientes_workbook,
)


def test_missing_required_columns_lists_absent_columns() -> None:
    missing = missing_required_columns(["RAZÃO SOCIAL"])

    assert "E-Mail 1" in missing
    assert "ArquivoAnexo" in missing


def test_read_clientes_workbook_parses_rows_and_preserves_line_numbers(tmp_path: Path) -> None:
    path = tmp_path / "clientes.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Clientes"
    worksheet.append(
        [
            " RAZÃO SOCIAL ",
            "Proprietário/Dirigente",
            "E-Mail 1",
            "Valor fev26",
            "Dia de Pagamento",
            "PIX",
            "ArquivoAnexo",
        ]
    )
    worksheet.append(["", "", "", "", "", "", ""])
    worksheet.append(["Empresa", "João", "joao@email.com", 300, 10, "pix", "arquivo.pdf"])
    workbook.save(path)

    data = read_clientes_workbook(path)

    assert data.headers[0] == "RAZÃO SOCIAL"
    assert len(data.rows) == 1
    assert data.rows[0].row_number == 3
    assert data.rows[0].values["RAZÃO SOCIAL"] == "Empresa"


def test_read_clientes_workbook_requires_clientes_sheet(tmp_path: Path) -> None:
    path = tmp_path / "clientes.xlsx"
    workbook = Workbook()
    workbook.active.title = "Outra"
    workbook.save(path)

    with pytest.raises(ExcelReaderError, match="Clientes"):
        read_clientes_workbook(path)
