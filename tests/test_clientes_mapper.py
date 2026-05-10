from __future__ import annotations

from pathlib import Path

from mailmerge_assistant.clientes_mapper import apply_test_mode, map_row_to_validation_result
from mailmerge_assistant.models import SpreadsheetRow


def _row(tmp_path: Path, **overrides: object) -> SpreadsheetRow:
    attachment = tmp_path / "anexo.pdf"
    attachment.write_text("pdf", encoding="utf-8")
    values: dict[str, object] = {
        "RAZÃO SOCIAL": "A F SERVICOS DE TRANSPORTES LTDA",
        "Proprietário/Dirigente": "Márcio",
        "E-Mail 1": "cliente@email.com",
        "E-Mail 2": "financeiro@email.com",
        "E-Mail 3": "",
        "Valor fev26": 390.5,
        "Dia de Pagamento": 10,
        "PIX": "pix@email.com",
        "ArquivoAnexo": str(attachment),
    }
    values.update(overrides)
    return SpreadsheetRow(row_number=8, values=values)


def test_map_row_combines_emails_and_generates_subject_body(tmp_path: Path) -> None:
    result = map_row_to_validation_result(_row(tmp_path))

    assert result.is_valid
    assert result.draft is not None
    assert result.draft.to == "cliente@email.com; financeiro@email.com"
    assert result.draft.subject == "Honorários contábeis - A F SERVICOS DE TRANSPORTES LTDA"
    assert "R$ 390,50" in result.draft.body
    assert "Márcio" in result.draft.body
    assert result.draft.attachments == []


def test_map_row_does_not_require_attachment_column(tmp_path: Path) -> None:
    row = _row(tmp_path)
    row.values.pop("ArquivoAnexo")

    result = map_row_to_validation_result(row)

    assert result.is_valid
    assert result.draft is not None
    assert result.draft.attachments == []


def test_map_row_blocks_excel_error_in_required_field(tmp_path: Path) -> None:
    result = map_row_to_validation_result(_row(tmp_path, **{"Dia de Pagamento": "#REF!"}))

    assert not result.is_valid
    assert "erro na planilha" in result.message


def test_map_row_blocks_invalid_email(tmp_path: Path) -> None:
    result = map_row_to_validation_result(_row(tmp_path, **{"E-Mail 1": "cliente@email"}))

    assert not result.is_valid
    assert "inválido" in result.message


def test_apply_test_mode_replaces_recipient_and_prefixes_subject(tmp_path: Path) -> None:
    result = map_row_to_validation_result(_row(tmp_path))
    assert result.draft is not None

    test_draft = apply_test_mode(result.draft, "teste@email.com")

    assert test_draft.to == "teste@email.com"
    assert test_draft.subject.startswith("[TESTE]")
    assert "Destinatário original: cliente@email.com; financeiro@email.com" in test_draft.body
