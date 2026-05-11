from __future__ import annotations

import base64
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from mailmerge_assistant.clientes_mapper import map_row_to_validation_result  # noqa: E402
from mailmerge_assistant.excel_reader import read_clientes_workbook  # noqa: E402
from mailmerge_assistant.outlook_client import LOGO_PATH, plain_text_to_html_email  # noqa: E402


def main() -> None:
    sample_path = PROJECT_ROOT / "samples" / "Cadastro de Empresas Contabilizuum - maio26.xlsx"
    workbook = read_clientes_workbook(sample_path)

    for row in workbook.rows:
        result = map_row_to_validation_result(row)
        if result.is_valid and result.draft is not None:
            logo_src = _logo_data_uri() if LOGO_PATH.exists() else None
            html = plain_text_to_html_email(result.draft.body, logo_src=logo_src)
            output_path = PROJECT_ROOT / "reports" / "email_preview.html"
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(html, encoding="utf-8")
            print(output_path)
            return

    raise RuntimeError("Nenhuma linha válida encontrada para gerar preview.")


def _logo_data_uri() -> str:
    encoded = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
    return f"data:image/jpeg;base64,{encoded}"


if __name__ == "__main__":
    main()
