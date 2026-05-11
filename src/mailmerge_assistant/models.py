from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

ReportStatus = Literal[
    "OK",
    "ERRO",
    "RASCUNHO_CRIADO",
    "IGNORADO",
    "ENVIADO",
    "FALHA_ENVIO",
]


@dataclass(frozen=True)
class SpreadsheetRow:
    row_number: int
    values: dict[str, Any]


@dataclass(frozen=True)
class EmailDraft:
    row_number: int
    razao_social: str
    proprietario: str
    to: str
    subject: str
    body: str
    valor: str
    dia_pagamento: str
    attachments: list[Path]
    original_to: str | None = None


@dataclass(frozen=True)
class RowValidationResult:
    row: SpreadsheetRow
    is_valid: bool
    message: str
    draft: EmailDraft | None = None


@dataclass(frozen=True)
class ValidationSummary:
    total_rows: int
    valid_rows: int
    invalid_rows: int


@dataclass(frozen=True)
class ValidationRunResult:
    rows: list[RowValidationResult]
    report_path: Path

    @property
    def summary(self) -> ValidationSummary:
        valid = sum(1 for row in self.rows if row.is_valid)
        return ValidationSummary(
            total_rows=len(self.rows),
            valid_rows=valid,
            invalid_rows=len(self.rows) - valid,
        )


@dataclass(frozen=True)
class DraftCreationResult:
    row_results: list[RowValidationResult]
    created_count: int
    report_path: Path


@dataclass(frozen=True)
class SendEmailsResult:
    row_results: list[RowValidationResult]
    sent_count: int
    report_path: Path


@dataclass(frozen=True)
class ReportRow:
    linha: int
    razao_social: str = ""
    proprietario: str = ""
    emails: str = ""
    assunto: str = ""
    valor: str = ""
    dia_pagamento: str = ""
    anexo: str = ""
    status: ReportStatus = "IGNORADO"
    mensagem: str = ""


@dataclass
class WorkbookData:
    headers: list[str] = field(default_factory=list)
    rows: list[SpreadsheetRow] = field(default_factory=list)
