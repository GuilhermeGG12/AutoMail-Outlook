from __future__ import annotations

from pathlib import Path

from mailmerge_assistant.controller import MailMergeController
from mailmerge_assistant.models import EmailDraft, RowValidationResult, SpreadsheetRow


class FakeOutlookClient:
    def __init__(self) -> None:
        self.created: list[EmailDraft] = []

    def create_draft(self, draft: object) -> None:
        assert isinstance(draft, EmailDraft)
        self.created.append(draft)


def test_controller_creates_drafts_with_fake_outlook_in_test_mode(tmp_path: Path) -> None:
    attachment = tmp_path / "anexo.pdf"
    attachment.write_text("pdf", encoding="utf-8")
    draft = EmailDraft(
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
