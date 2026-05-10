from __future__ import annotations

import pytest

from mailmerge_assistant.template_engine import TemplateRenderError, render_template


def test_render_template_replaces_placeholders() -> None:
    assert render_template("Olá {nome}", {"nome": "Maria"}) == "Olá Maria"


def test_render_template_reports_missing_placeholders() -> None:
    with pytest.raises(TemplateRenderError):
        render_template("Olá {nome}", {})
