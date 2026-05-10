from __future__ import annotations

from mailmerge_assistant.models import EmailDraft


class OutlookDraftError(RuntimeError):
    """Friendly Outlook error."""


class OutlookClient:
    def __init__(self) -> None:
        try:
            import win32com.client
        except ImportError as exc:
            raise OutlookDraftError(
                "Não foi possível carregar a integração com Outlook. "
                "Instale o pywin32 e execute no Windows com Outlook desktop instalado."
            ) from exc
        self._win32_client = win32com.client

    def create_draft(self, draft: EmailDraft) -> None:
        try:
            outlook = self._win32_client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = draft.to
            mail.Subject = draft.subject
            mail.Body = draft.body
            for attachment in draft.attachments:
                mail.Attachments.Add(str(attachment))
            mail.Save()
        except Exception as exc:
            raise OutlookDraftError(
                f"Não foi possível criar o rascunho da linha {draft.row_number} no Outlook."
            ) from exc
