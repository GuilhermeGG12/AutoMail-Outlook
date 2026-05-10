from __future__ import annotations

from mailmerge_assistant.outlook_client import plain_text_to_html_email


def test_plain_text_to_html_email_preserves_paragraphs_and_escapes_html() -> None:
    html = plain_text_to_html_email("Olá João,\nLinha 2\n\nPIX <teste>&valor")

    assert "<p>Olá João,<br>Linha 2</p>" in html
    assert "<p>PIX &lt;teste&gt;&amp;valor</p>" in html
    assert "font-family: Calibri" in html
