from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from mailmerge_assistant.config import EXCEL_ERROR_VALUES

EMAIL_RE = re.compile(r"^[^@\s;]+@[^@\s;]+\.[^@\s;]+$")


def normalize_header(value: Any) -> str:
    return "" if value is None else str(value).strip()


def stringify_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def is_excel_error(value: Any) -> bool:
    return stringify_cell(value).upper() in EXCEL_ERROR_VALUES


def is_blank(value: Any) -> bool:
    return stringify_cell(value) == ""


def split_semicolon_values(value: Any) -> list[str]:
    text = stringify_cell(value)
    if not text:
        return []
    return [part.strip() for part in text.split(";") if part.strip()]


def validate_email_address(email: str) -> bool:
    return bool(EMAIL_RE.fullmatch(email.strip()))


def parse_email_list(value: Any, *, required: bool = False) -> tuple[list[str], str | None]:
    if is_excel_error(value):
        return [], "o campo de e-mail está com erro na planilha. Corrija no Excel."
    emails = split_semicolon_values(value)
    if required and not emails:
        return [], 'o campo "E-Mail 1" é obrigatório.'
    for email in emails:
        if not validate_email_address(email):
            return [], f'o e-mail "{email}" é inválido.'
    return emails, None


def parse_brl_currency(value: Any) -> tuple[str | None, str | None]:
    if is_excel_error(value):
        return None, 'o campo "Valor fev26" está com erro na planilha. Corrija no Excel.'
    if is_blank(value):
        return None, 'o campo "Valor fev26" é obrigatório.'
    if isinstance(value, int | float | Decimal):
        amount = Decimal(str(value))
    else:
        text = stringify_cell(value).replace("R$", "").replace(" ", "")
        if "," in text and "." in text:
            if text.rfind(".") > text.rfind(","):
                text = text.replace(",", "")
            else:
                text = text.replace(".", "").replace(",", ".")
        elif "," in text:
            text = text.replace(",", ".")
        try:
            amount = Decimal(text)
        except InvalidOperation:
            return None, 'o campo "Valor fev26" não parece ser um valor numérico.'
    formatted = f"{amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {formatted}", None


def validate_required_value(value: Any, field_name: str) -> str | None:
    if is_excel_error(value):
        return (
            f'o campo "{field_name}" está com erro na planilha. '
            "Corrija no Excel antes de gerar o e-mail."
        )
    if is_blank(value):
        return f'o campo "{field_name}" é obrigatório.'
    return None


def validate_attachment_paths(value: Any) -> tuple[list[Path], str | None]:
    if is_excel_error(value):
        return [], 'o campo "ArquivoAnexo" está com erro na planilha. Corrija no Excel.'
    paths = split_semicolon_values(value)
    if not paths:
        return [], 'o campo "ArquivoAnexo" é obrigatório.'
    valid_paths: list[Path] = []
    for raw_path in paths:
        if raw_path.upper() in EXCEL_ERROR_VALUES:
            return [], 'o campo "ArquivoAnexo" contém um erro da planilha.'
        path = Path(raw_path).expanduser()
        if not path.exists():
            return [], (
                f'Anexo não encontrado: "{_friendly_path_name(path)}". '
                "Verifique se o arquivo existe neste computador ou ajuste a coluna ArquivoAnexo."
            )
        if path.is_dir():
            return [], (
                f'O anexo "{_friendly_path_name(path)}" aponta para uma pasta, '
                "não para um arquivo."
            )
        valid_paths.append(path)
    return valid_paths, None


def _friendly_path_name(path: Path) -> str:
    name = path.name or str(path)
    if len(name) <= 80:
        return name
    suffix = path.suffix
    stem_limit = max(20, 77 - len(suffix))
    return f"{name[:stem_limit]}...{suffix}"
