from __future__ import annotations

import os
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

import customtkinter as ctk

from mailmerge_assistant.config import APP_NAME, REPORTS_DIR
from mailmerge_assistant.controller import MailMergeController
from mailmerge_assistant.models import RowValidationResult


class MainWindow(ctk.CTk):  # type: ignore[misc]
    def __init__(self, controller: MailMergeController | None = None) -> None:
        super().__init__()
        self.controller = controller or MailMergeController()
        self.selected_path: Path | None = None
        self.title(APP_NAME)
        self.geometry("1180x720")
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self._build_ui()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        top = ctk.CTkFrame(self)
        top.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 8))
        top.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(top, text="Selecionar planilha", command=self._select_file).grid(
            row=0, column=0, padx=8, pady=8
        )
        self.path_label = ctk.CTkLabel(top, text="Nenhuma planilha selecionada", anchor="w")
        self.path_label.grid(row=0, column=1, sticky="ew", padx=8)
        ctk.CTkButton(top, text="Validar clientes", command=self._validate).grid(
            row=0, column=2, padx=8, pady=8
        )

        self.summary_label = ctk.CTkLabel(
            self,
            text="Total: 0 | Válidos: 0 | Inválidos: 0",
            anchor="w",
        )
        self.summary_label.grid(row=1, column=0, sticky="ew", padx=16)

        self.table = ttk.Treeview(
            self,
            columns=(
                "status",
                "linha",
                "razao",
                "proprietario",
                "emails",
                "valor",
                "dia",
                "anexo",
                "mensagem",
            ),
            show="headings",
        )
        headings = {
            "status": "Status",
            "linha": "Linha",
            "razao": "Razão Social",
            "proprietario": "Proprietário/Dirigente",
            "emails": "Emails",
            "valor": "Valor",
            "dia": "Dia de Pagamento",
            "anexo": "Anexo",
            "mensagem": "Mensagem",
        }
        widths = {
            "status": 90,
            "linha": 60,
            "razao": 210,
            "proprietario": 170,
            "emails": 210,
            "valor": 110,
            "dia": 130,
            "anexo": 190,
            "mensagem": 280,
        }
        for column, heading in headings.items():
            self.table.heading(column, text=heading)
            self.table.column(column, width=widths[column], minwidth=60)
        self.table.grid(row=2, column=0, sticky="nsew", padx=12, pady=8)

        bottom = ctk.CTkFrame(self)
        bottom.grid(row=3, column=0, sticky="ew", padx=12, pady=(8, 12))
        bottom.grid_columnconfigure(2, weight=1)

        self.test_mode_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(bottom, text="Modo de teste", variable=self.test_mode_var).grid(
            row=0, column=0, padx=8, pady=8
        )
        self.test_email_entry = ctk.CTkEntry(bottom, placeholder_text="E-mail de teste")
        self.test_email_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=8)
        ctk.CTkButton(
            bottom,
            text="Visualizar e-mail",
            command=self._preview_email,
        ).grid(row=0, column=3, padx=8, pady=8)
        ctk.CTkButton(
            bottom,
            text="Criar rascunhos no Outlook",
            command=self._create_drafts,
        ).grid(row=0, column=4, padx=8, pady=8)
        ctk.CTkButton(
            bottom,
            text="Enviar e-mails",
            fg_color="#b42318",
            hover_color="#8f1d14",
            command=self._send_emails,
        ).grid(row=0, column=5, padx=8, pady=8)
        ctk.CTkButton(
            bottom,
            text="Abrir pasta de relatórios",
            command=self._open_reports,
        ).grid(row=0, column=6, padx=8, pady=8)

    def _select_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="Selecionar planilha",
            filetypes=[("Planilhas Excel", "*.xlsx")],
        )
        if filename:
            self.selected_path = Path(filename)
            self.path_label.configure(text=str(self.selected_path))

    def _validate(self) -> None:
        if self.selected_path is None:
            messagebox.showwarning(APP_NAME, "Selecione uma planilha antes de validar.")
            return
        try:
            result = self.controller.validate_file(self.selected_path)
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        summary = result.summary
        self.summary_label.configure(
            text=(
                f"Total: {summary.total_rows} | "
                f"Válidos: {summary.valid_rows} | Inválidos: {summary.invalid_rows}"
            )
        )
        self._populate_table(result.rows)
        messagebox.showinfo(APP_NAME, f"Validação concluída. Relatório: {result.report_path}")

    def _populate_table(self, rows: list[RowValidationResult]) -> None:
        for item in self.table.get_children():
            self.table.delete(item)
        for result in rows:
            draft = result.draft
            self.table.insert(
                "",
                "end",
                values=(
                    "OK" if result.is_valid else "ERRO",
                    result.row.row_number,
                    draft.razao_social if draft else result.row.values.get("RAZÃO SOCIAL", ""),
                    (
                        draft.proprietario
                        if draft
                        else result.row.values.get("Proprietário/Dirigente", "")
                    ),
                    draft.to if draft else result.row.values.get("E-Mail 1", ""),
                    draft.valor if draft else result.row.values.get("Valor fev26", ""),
                    draft.dia_pagamento if draft else result.row.values.get("Dia de Pagamento", ""),
                    (
                        "; ".join(str(path) for path in draft.attachments)
                        if draft
                        else result.row.values.get("ArquivoAnexo", "")
                    ),
                    result.message,
                ),
            )

    def _create_drafts(self) -> None:
        valid_count = sum(1 for row in self.controller.last_validation if row.is_valid)
        if valid_count == 0:
            messagebox.showwarning(APP_NAME, "Não há linhas válidas para criar rascunhos.")
            return
        confirmed = messagebox.askyesno(
            APP_NAME,
            f"Serão criados {valid_count} rascunhos no Outlook. Deseja continuar?",
        )
        if not confirmed:
            return
        try:
            result = self.controller.create_outlook_drafts(
                test_mode=self.test_mode_var.get(),
                test_email=self.test_email_entry.get().strip(),
            )
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self._populate_table(result.row_results)
        messagebox.showinfo(
            APP_NAME,
            f"{result.created_count} rascunhos criados. Relatório: {result.report_path}",
        )

    def _preview_email(self) -> None:
        try:
            preview_path = self.controller.write_preview_html(
                test_mode=self.test_mode_var.get(),
                test_email=self.test_email_entry.get().strip(),
            )
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        os.startfile(str(preview_path.resolve()))

    def _send_emails(self) -> None:
        valid_count = sum(1 for row in self.controller.last_validation if row.is_valid)
        invalid_count = sum(1 for row in self.controller.last_validation if not row.is_valid)
        if valid_count == 0:
            messagebox.showwarning(APP_NAME, "Não há linhas válidas para enviar.")
            return
        if invalid_count:
            messagebox.showerror(
                APP_NAME,
                f"Existem {invalid_count} linhas inválidas. Corrija a planilha antes de enviar.",
            )
            return

        mode_text = "de TESTE" if self.test_mode_var.get() else "REAIS"
        confirmed = messagebox.askyesno(
            APP_NAME,
            (
                f"Serão enviados {valid_count} e-mails {mode_text} pelo Outlook.\n\n"
                "Essa ação usa o botão Enviar do Outlook. Deseja continuar?"
            ),
        )
        if not confirmed:
            return

        phrase = simpledialog.askstring(
            APP_NAME,
            f'Serão enviados {valid_count} e-mails {mode_text}. Digite "ENVIAR" para continuar:',
            parent=self,
        )
        if phrase is None:
            return

        try:
            result = self.controller.send_outlook_emails(
                confirmation_phrase=phrase,
                test_mode=self.test_mode_var.get(),
                test_email=self.test_email_entry.get().strip(),
            )
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self._populate_table(result.row_results)
        messagebox.showinfo(
            APP_NAME,
            f"{result.sent_count} e-mails enviados. Relatório: {result.report_path}",
        )

    def _open_reports(self) -> None:
        REPORTS_DIR.mkdir(parents=True, exist_ok=True)
        os.startfile(str(REPORTS_DIR.resolve()))
