from __future__ import annotations

from mailmerge_assistant.outlook_client import plain_text_to_html_email


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
