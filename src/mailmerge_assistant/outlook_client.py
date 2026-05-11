from __future__ import annotations

import re
from html import escape
from pathlib import Path
from typing import Any

from mailmerge_assistant.models import EmailDraft

LOGO_CONTENT_ID = "mailmerge-assistant-company-logo"
LOGO_PATH = Path(__file__).with_name("assets") / "company_logo.jpeg"
STRONG_MARKUP_RE = re.compile(r"\*\*(.+?)\*\*")


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
            mail = self._create_mail_item(draft)
            mail.Save()
        except OutlookDraftError:
            raise
        except Exception as exc:
            raise OutlookDraftError(
                f"Não foi possível criar o rascunho da linha {draft.row_number} no Outlook."
            ) from exc

    def send_email(self, draft: EmailDraft) -> None:
        try:
            mail = self._create_mail_item(draft)
            mail.Send()
        except OutlookDraftError:
            raise
        except Exception as exc:
            raise OutlookDraftError(
                f"Não foi possível enviar o e-mail da linha {draft.row_number} pelo Outlook."
            ) from exc

    def refresh_drafts_folder(self, *, max_items: int = 60) -> None:
        try:
            outlook = self._win32_client.Dispatch("Outlook.Application")
            drafts = outlook.Session.GetDefaultFolder(16)
            items = drafts.Items
            items.Sort("[LastModificationTime]", True)
            limit = min(max_items, int(items.Count))
            for index in range(1, limit + 1):
                item = items.Item(index)
                _ = item.Subject
                _ = item.To
                _ = item.LastModificationTime
        except Exception:
            return

    def _create_mail_item(self, draft: EmailDraft) -> Any:
        outlook = self._win32_client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = draft.to
        if not mail.Recipients.ResolveAll():
            raise OutlookDraftError(
                f'Não foi possível validar o destinatário "{draft.to}" no Outlook.'
            )
        mail.Subject = draft.subject
        mail.Importance = 2  # olImportanceHigh
        mail.ReadReceiptRequested = True
        mail.BodyFormat = 2  # olFormatHTML
        logo_src = f"cid:{LOGO_CONTENT_ID}" if LOGO_PATH.exists() else None
        mail.HTMLBody = plain_text_to_html_email(draft.body, logo_src=logo_src)
        if logo_src is not None:
            logo = mail.Attachments.Add(str(LOGO_PATH))
            logo.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                LOGO_CONTENT_ID,
            )
            logo.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
                True,
            )
        return mail


def plain_text_to_html_email(text: str, *, logo_src: str | None = None) -> str:
    blocks = [block.strip() for block in text.replace("\r\n", "\n").split("\n\n") if block.strip()]
    rendered_blocks = []
    for block in blocks:
        lines = [_render_inline_markup(line.strip()) for line in block.split("\n")]
        rendered_blocks.append(f"<p>{'<br>'.join(lines)}</p>")

    if logo_src is not None:
        rendered_blocks.append(
            '<p class="logo">'
            f'<img src="{escape(logo_src, quote=True)}" '
            'alt="ContabiliZuum.com Contabilidade Digital">'
            "</p>"
        )

    body = "\n".join(rendered_blocks)
    return f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body {{
      font-family: Calibri, Arial, sans-serif;
      font-size: 11pt;
      color: #1f1f1f;
      line-height: 1.35;
    }}
    p {{
      margin: 0 0 12px 0;
    }}
    .logo {{
      margin-top: 18px;
    }}
    .logo img {{
      width: 300px;
      height: auto;
      display: block;
    }}
  </style>
</head>
<body>
{body}
</body>
</html>"""


def _render_inline_markup(text: str) -> str:
    pieces: list[str] = []
    last_index = 0
    for match in STRONG_MARKUP_RE.finditer(text):
        pieces.append(escape(text[last_index : match.start()]))
        pieces.append(f"<strong>{escape(match.group(1))}</strong>")
        last_index = match.end()
    pieces.append(escape(text[last_index:]))
    return "".join(pieces)
