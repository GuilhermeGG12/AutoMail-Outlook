from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook


def main() -> None:
    samples_dir = Path("samples")
    samples_dir.mkdir(exist_ok=True)
    attachment = samples_dir / "boleto_exemplo.pdf"
    if not attachment.exists():
        attachment.write_bytes(b"%PDF-1.4\n% sample\n")

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Clientes"
    worksheet.append(
        [
            "RAZÃO SOCIAL",
            "Proprietário/Dirigente",
            "E-Mail 1",
            "E-Mail 2",
            "E-Mail 3",
            "Valor fev26",
            "Dia de Pagamento",
            "PIX",
            "ArquivoAnexo",
        ]
    )
    worksheet.append(
        [
            "EMPRESA EXEMPLO LTDA",
            "Maria",
            "cliente@example.com",
            "financeiro@example.com",
            "",
            390.5,
            10,
            "pix@example.com",
            str(attachment.resolve()),
        ]
    )
    workbook.save(samples_dir / "sample_clientes.xlsx")


if __name__ == "__main__":
    main()
