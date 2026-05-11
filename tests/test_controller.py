from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import load_workbook

from mailmerge_assistant.controller import MailMergeController
from mailmerge_assistant.models import EmailDraft, RowValidationResult, SpreadsheetRow


class FakeOutlookClient:
    def __init__(self) -> None:
        self.created: list[EmailDraft] = []
        self.sent: list[EmailDraft] = []
        self.refreshed_max_items: int | None = None

    def create_draft(self, draft: object) -> None:
        assert isinstance(draft, EmailDraft)
        self.created.append(draft)

    def send_email(self, draft: object) -> None:
        assert isinstance(draft, EmailDraft)
        self.sent.append(draft)

    def refresh_drafts_folder(self, *, max_items: int = 60) -> None:
        self.refreshed_max_items = max_items


def test_controller_creates_drafts_with_fake_outlook_in_test_mode(tmp_path: Path) -> None:
    draft = _draft(tmp_path)
    fake_client = FakeOutlookClient()
    controller = MailMergeController(outlook_client=fake_client)
    controller._last_validation = [  # noqa: SLF001
        RowValidationResult(
            row=SpreadsheetRow(row_number=2, values={}),
            is_valid=True,
            message="OK",
            draft=draft,
        )
    ]

    result = controller.create_outlook_drafts(test_mode=True, test_email="teste@email.com")

    assert result.created_count == 1
    assert fake_client.created[0].to == "teste@email.com"
    assert fake_client.created[0].subject == "[TESTE] Honorários contábeis - Empresa"
    assert fake_client.created[0].original_to == "cliente@email.com"
    assert fake_client.refreshed_max_items == 60


def test_controller_sends_emails_only_with_confirmation_in_test_mode(tmp_path: Path) -> None:
    draft = _draft(tmp_path)
    fake_client = FakeOutlookClient()
    controller = MailMergeController(outlook_client=fake_client)
    controller._last_validation = [  # noqa: SLF001
        RowValidationResult(
            row=SpreadsheetRow(row_number=2, values={}),
            is_valid=True,
            message="OK",
            draft=draft,
        )
    ]

    result = controller.send_outlook_emails(
        confirmation_phrase="ENVIAR",
        test_mode=True,
        test_email="teste@email.com",
    )

    assert result.sent_count == 1
    assert fake_client.sent[0].to == "teste@email.com"
    assert fake_client.sent[0].subject == "[TESTE] Honorários contábeis - Empresa"
    assert fake_client.sent[0].original_to == "cliente@email.com"

    workbook = load_workbook(result.report_path)
    worksheet = workbook.active
    assert worksheet["J2"].value == "ENVIADO"


def test_controller_blocks_send_without_confirmation(tmp_path: Path) -> None:
    fake_client = FakeOutlookClient()
    controller = MailMergeController(outlook_client=fake_client)
    controller._last_validation = [  # noqa: SLF001
        RowValidationResult(
            row=SpreadsheetRow(row_number=2, values={}),
            is_valid=True,
            message="OK",
            draft=_draft(tmp_path),
        )
    ]

    with pytest.raises(ValueError, match="ENVIAR"):
        controller.send_outlook_emails(confirmation_phrase="sim")

    assert fake_client.sent == []


def test_controller_writes_preview_html(tmp_path: Path) -> None:
    controller = MailMergeController(outlook_client=FakeOutlookClient())
    controller._last_validation = [  # noqa: SLF001
        RowValidationResult(
            row=SpreadsheetRow(row_number=2, values={}),
            is_valid=True,
            message="OK",
            draft=_draft(tmp_path),
        )
    ]

    preview_path = controller.write_preview_html()

    assert preview_path.exists()
    assert "<html>" in preview_path.read_text(encoding="utf-8")


def test_controller_blocks_send_when_any_row_is_invalid(tmp_path: Path) -> None:
    fake_client = FakeOutlookClient()
    controller = MailMergeController(outlook_client=fake_client)
    controller._last_validation = [  # noqa: SLF001
        RowValidationResult(
            row=SpreadsheetRow(row_number=2, values={}),
            is_valid=True,
            message="OK",
            draft=_draft(tmp_path),
        ),
        RowValidationResult(
            row=SpreadsheetRow(row_number=3, values={}),
            is_valid=False,
            message="Linha 3: erro",
            draft=None,
        ),
    ]

    with pytest.raises(ValueError, match="linhas inválidas"):
        controller.send_outlook_emails(confirmation_phrase="ENVIAR")

    assert fake_client.sent == []


def _draft(tmp_path: Path) -> EmailDraft:
    attachment = tmp_path / "anexo.pdf"
    attachment.write_text("pdf", encoding="utf-8")
    return EmailDraft(
        row_number=2,
        razao_social="Empresa",
        proprietario="Maria",
        to="cliente@email.com",
        subject="Honorários contábeis - Empresa",
        body="Corpo",
        valor="R$ 300,00",
        dia_pagamento="10",
        attachments=[attachment],
    )
