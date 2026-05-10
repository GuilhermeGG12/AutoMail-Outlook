from __future__ import annotations

from html import escape

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
            mail.BodyFormat = 2  # olFormatHTML
            mail.HTMLBody = plain_text_to_html_email(draft.body)
            mail.Save()
        except Exception as exc:
            raise OutlookDraftError(
                f"Não foi possível criar o rascunho da linha {draft.row_number} no Outlook."
            ) from exc


def plain_text_to_html_email(text: str) -> str:
    blocks = [block.strip() for block in text.replace("\r\n", "\n").split("\n\n") if block.strip()]
    rendered_blocks = []
    for block in blocks:
        lines = [escape(line.strip()) for line in block.split("\n")]
        rendered_blocks.append(f"<p>{'<br>'.join(lines)}</p>")

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
  </style>
</head>
<body>
{body}
</body>
</html>"""
