from __future__ import annotations

from pathlib import Path

import pytest

from mailmerge_assistant.models import EmailDraft
from mailmerge_assistant.outlook_client import (
    OutlookClient,
    OutlookDraftError,
    plain_text_to_html_email,
)


def test_plain_text_to_html_email_preserves_paragraphs_and_escapes_html() -> None:
    html = plain_text_to_html_email("Olá João,\nLinha 2\n\nPIX <teste>&valor")

    assert "<p>Olá João,<br>Linha 2</p>" in html
    assert "<p>PIX &lt;teste&gt;&amp;valor</p>" in html
    assert "font-family: Calibri" in html


def test_plain_text_to_html_email_can_render_logo() -> None:
    html = plain_text_to_html_email("Corpo", logo_src="cid:logo&teste")

    assert '<p class="logo"><img src="cid:logo&amp;teste"' in html
    assert "ContabiliZuum.com Contabilidade Digital" in html


def test_plain_text_to_html_email_renders_strong_markup_safely() -> None:
    html = plain_text_to_html_email("Olá **João <teste>**")

    assert "<p>Olá <strong>João &lt;teste&gt;</strong></p>" in html


def test_create_draft_resolves_recipients_before_saving(tmp_path: Path) -> None:
    fake_win32 = FakeWin32Client(resolve_result=True)
    client = OutlookClient.__new__(OutlookClient)
    client._win32_client = fake_win32  # noqa: SLF001

    client.create_draft(_draft(tmp_path))

    mail = fake_win32.outlook.mail
    assert mail.Recipients.resolve_all_called
    assert mail.Importance == 2
    assert mail.ReadReceiptRequested
    assert mail.saved


def test_create_draft_blocks_unresolved_recipient(tmp_path: Path) -> None:
    fake_win32 = FakeWin32Client(resolve_result=False)
    client = OutlookClient.__new__(OutlookClient)
    client._win32_client = fake_win32  # noqa: SLF001

    with pytest.raises(OutlookDraftError, match="destinat"):
        client.create_draft(_draft(tmp_path))

    assert not fake_win32.outlook.mail.saved


class FakeRecipients:
    def __init__(self, resolve_result: bool) -> None:
        self.resolve_result = resolve_result
        self.resolve_all_called = False

    def ResolveAll(self) -> bool:
        self.resolve_all_called = True
        return self.resolve_result


class FakeAttachments:
    def Add(self, _path: str) -> object:
        return FakeAttachment()


class FakeAttachment:
    @property
    def PropertyAccessor(self) -> object:
        return self

    def SetProperty(self, _name: str, _value: object) -> None:
        return None


class FakeMail:
    def __init__(self, resolve_result: bool) -> None:
        self.To = ""
        self.Subject = ""
        self.Importance = 1
        self.ReadReceiptRequested = False
        self.BodyFormat = 0
        self.HTMLBody = ""
        self.Recipients = FakeRecipients(resolve_result)
        self.Attachments = FakeAttachments()
        self.saved = False

    def Save(self) -> None:
        self.saved = True

    def Send(self) -> None:
        return None


class FakeOutlookApplication:
    def __init__(self, resolve_result: bool) -> None:
        self.mail = FakeMail(resolve_result)

    def CreateItem(self, _kind: int) -> FakeMail:
        return self.mail


class FakeWin32Client:
    def __init__(self, resolve_result: bool) -> None:
        self.outlook = FakeOutlookApplication(resolve_result)

    def Dispatch(self, _name: str) -> FakeOutlookApplication:
        return self.outlook


def _draft(_tmp_path: Path) -> EmailDraft:
    return EmailDraft(
        row_number=2,
        razao_social="Empresa",
        proprietario="Maria",
        to="cliente@email.com",
        subject="Honorários contábeis - Empresa",
        body="Corpo",
        valor="R$ 300,00",
        dia_pagamento="10",
        attachments=[],
    )
