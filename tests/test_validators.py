from __future__ import annotations

from pathlib import Path

from mailmerge_assistant.validators import (
    is_excel_error,
    parse_brl_currency,
    parse_email_list,
    validate_attachment_paths,
)


def test_parse_email_list_accepts_semicolon_values() -> None:
    emails, error = parse_email_list("cliente@email.com; financeiro@email.com", required=True)

    assert error is None
    assert emails == ["cliente@email.com", "financeiro@email.com"]


def test_parse_email_list_blocks_invalid_email() -> None:
    emails, error = parse_email_list("cliente@email", required=True)

    assert emails == []
    assert error is not None


def test_parse_brl_currency_formats_values() -> None:
    assert parse_brl_currency(300)[0] == "R$ 300,00"
    assert parse_brl_currency(390.5)[0] == "R$ 390,50"
    assert parse_brl_currency("1,500.75")[0] == "R$ 1.500,75"


def test_parse_brl_currency_blocks_invalid_value() -> None:
    value, error = parse_brl_currency("abc")

    assert value is None
    assert error is not None


def test_excel_error_detection() -> None:
    assert is_excel_error("#REF!")
    assert is_excel_error("#N/A")


def test_validate_attachment_paths(tmp_path: Path) -> None:
    attachment = tmp_path / "arquivo.pdf"
    attachment.write_text("x", encoding="utf-8")

    paths, error = validate_attachment_paths(str(attachment))

    assert error is None
    assert paths == [attachment]


def test_validate_attachment_paths_blocks_missing_file() -> None:
    paths, error = validate_attachment_paths("C:/arquivo/inexistente.pdf")

    assert paths == []
    assert error is not None
