from __future__ import annotations

from typing import Any

from mailmerge_assistant.config import EMAIL_BODY_TEMPLATE, EMAIL_SUBJECT_TEMPLATE, REQUIRED_COLUMNS
from mailmerge_assistant.models import EmailDraft, RowValidationResult, SpreadsheetRow
from mailmerge_assistant.template_engine import TemplateRenderError, render_template
from mailmerge_assistant.validators import (
    parse_brl_currency,
    parse_email_list,
    stringify_cell,
    validate_required_value,
)


def map_row_to_validation_result(row: SpreadsheetRow) -> RowValidationResult:
    try:
        draft = map_row_to_draft(row)
    except ValueError as exc:
        return RowValidationResult(
            row=row, is_valid=False, message=f"Linha {row.row_number}: {exc}"
        )
    return RowValidationResult(row=row, is_valid=True, message="OK", draft=draft)


def map_row_to_draft(row: SpreadsheetRow) -> EmailDraft:
    values = row.values
    for field_name in REQUIRED_COLUMNS:
        error = validate_required_value(values.get(field_name), field_name)
        if error:
            raise ValueError(error)

    email_1, error = parse_email_list(values.get("E-Mail 1"), required=True)
    if error:
        raise ValueError(error)
    email_2, error = parse_email_list(values.get("E-Mail 2"), required=False)
    if error:
        raise ValueError(error)
    email_3, error = parse_email_list(values.get("E-Mail 3"), required=False)
    if error:
        raise ValueError(error)
    recipients = [*email_1, *email_2, *email_3]
    to = "; ".join(recipients)

    valor, error = parse_brl_currency(values.get("Valor fev26"))
    if error or valor is None:
        raise ValueError(error or 'o campo "Valor fev26" é inválido.')

    template_values = _template_values(values, valor)
    try:
        subject = render_template(EMAIL_SUBJECT_TEMPLATE, template_values).strip()
        body = render_template(EMAIL_BODY_TEMPLATE, template_values).strip()
    except TemplateRenderError as exc:
        raise ValueError(str(exc)) from exc
    if not subject:
        raise ValueError("o assunto do e-mail ficou vazio.")
    if not body:
        raise ValueError("o corpo do e-mail ficou vazio.")

    return EmailDraft(
        row_number=row.row_number,
        razao_social=template_values["RAZÃO SOCIAL"],
        proprietario=template_values["Proprietário/Dirigente"],
        to=to,
        subject=subject,
        body=body,
        valor=valor,
        dia_pagamento=template_values["Dia de Pagamento"],
        attachments=[],
    )


def apply_test_mode(draft: EmailDraft, test_email: str) -> EmailDraft:
    return EmailDraft(
        row_number=draft.row_number,
        razao_social=draft.razao_social,
        proprietario=draft.proprietario,
        to=test_email,
        subject=f"[TESTE] {draft.subject}",
        body=f"MODO DE TESTE\nDestinatário original: {draft.to}\n\n---\n\n{draft.body}",
        valor=draft.valor,
        dia_pagamento=draft.dia_pagamento,
        attachments=draft.attachments,
        original_to=draft.to,
    )


def _template_values(values: dict[str, Any], valor: str) -> dict[str, str]:
    return {
        "Proprietário/Dirigente": stringify_cell(values.get("Proprietário/Dirigente")),
        "RAZÃO SOCIAL": stringify_cell(values.get("RAZÃO SOCIAL")),
        "Valor fev26": valor,
        "Dia de Pagamento": stringify_cell(values.get("Dia de Pagamento")),
        "PIX": stringify_cell(values.get("PIX")),
    }
