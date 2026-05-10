from __future__ import annotations

from pathlib import Path
from typing import Protocol, cast

from mailmerge_assistant.clientes_mapper import apply_test_mode, map_row_to_validation_result
from mailmerge_assistant.excel_reader import read_clientes_workbook
from mailmerge_assistant.models import (
    DraftCreationResult,
    ReportRow,
    ReportStatus,
    RowValidationResult,
    ValidationRunResult,
)
from mailmerge_assistant.outlook_client import OutlookClient
from mailmerge_assistant.report_writer import write_report
from mailmerge_assistant.validators import parse_email_list


class DraftClient(Protocol):
    def create_draft(self, draft: object) -> None: ...


class MailMergeController:
    def __init__(self, outlook_client: DraftClient | None = None) -> None:
        self._outlook_client = outlook_client
        self._last_validation: list[RowValidationResult] = []

    def validate_file(self, spreadsheet_path: str | Path) -> ValidationRunResult:
        workbook_data = read_clientes_workbook(spreadsheet_path)
        results = [map_row_to_validation_result(row) for row in workbook_data.rows]
        self._last_validation = results
        report_path = write_report([_validation_to_report_row(result) for result in results])
        return ValidationRunResult(rows=results, report_path=report_path)

    def create_outlook_drafts(
        self,
        *,
        test_mode: bool = False,
        test_email: str = "",
    ) -> DraftCreationResult:
        if not self._last_validation:
            raise ValueError("Valide a planilha antes de criar rascunhos.")
        if test_mode:
            emails, error = parse_email_list(test_email, required=True)
            if error or len(emails) != 1:
                raise ValueError("Informe um único e-mail de teste válido.")
            test_email = emails[0]

        client = self._outlook_client or OutlookClient()
        created_count = 0
        updated_results: list[RowValidationResult] = []
        report_rows: list[ReportRow] = []
        for result in self._last_validation:
            if not result.is_valid or result.draft is None:
                updated_results.append(result)
                report_rows.append(_validation_to_report_row(result, status="IGNORADO"))
                continue
            draft = apply_test_mode(result.draft, test_email) if test_mode else result.draft
            try:
                client.create_draft(draft)
            except Exception as exc:
                error_result = RowValidationResult(
                    row=result.row,
                    is_valid=False,
                    message=str(exc),
                    draft=result.draft,
                )
                updated_results.append(error_result)
                report_rows.append(_validation_to_report_row(error_result, status="ERRO"))
                continue
            created_count += 1
            created_result = RowValidationResult(
                row=result.row,
                is_valid=True,
                message="Rascunho criado no Outlook.",
                draft=draft,
            )
            updated_results.append(created_result)
            report_rows.append(_validation_to_report_row(created_result, status="RASCUNHO_CRIADO"))
        report_path = write_report(report_rows)
        return DraftCreationResult(
            row_results=updated_results,
            created_count=created_count,
            report_path=report_path,
        )

    @property
    def last_validation(self) -> list[RowValidationResult]:
        return self._last_validation


def _validation_to_report_row(
    result: RowValidationResult,
    *,
    status: str | None = None,
) -> ReportRow:
    draft = result.draft
    row = result.row
    return ReportRow(
        linha=row.row_number,
        razao_social=draft.razao_social if draft else str(row.values.get("RAZÃO SOCIAL") or ""),
        proprietario=(
            draft.proprietario if draft else str(row.values.get("Proprietário/Dirigente") or "")
        ),
        emails=draft.to if draft else str(row.values.get("E-Mail 1") or ""),
        assunto=draft.subject if draft else "",
        valor=draft.valor if draft else str(row.values.get("Valor fev26") or ""),
        dia_pagamento=(
            draft.dia_pagamento if draft else str(row.values.get("Dia de Pagamento") or "")
        ),
        anexo=(
            "; ".join(str(path) for path in draft.attachments)
            if draft
            else str(row.values.get("ArquivoAnexo") or "")
        ),
        status=cast(ReportStatus, status or ("OK" if result.is_valid else "ERRO")),
        mensagem=result.message,
    )
