from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile


class TemplateReaderError(ValueError):
    """Friendly template reader error."""


WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
MERGE_FIELD_PLACEHOLDERS = {
    "ProprietárioDirigente": "Proprietário/Dirigente",
    "ProprietarioDirigente": "Proprietário/Dirigente",
    "RAZÃO_SOCIAL": "RAZÃO SOCIAL",
    "RAZAO_SOCIAL": "RAZÃO SOCIAL",
    "Dia_de_Pagamento": "Dia de Pagamento",
    "Valor_fev26": "Valor fev26",
    "PIX": "PIX",
}


def read_email_body_template(path: str | Path) -> str:
    template_path = Path(path)
    if not template_path.exists():
        raise TemplateReaderError("O arquivo de modelo selecionado não foi encontrado.")
    if template_path.is_dir():
        raise TemplateReaderError("Selecione um arquivo de modelo, não uma pasta.")

    suffix = template_path.suffix.lower()
    if suffix == ".txt":
        return _read_txt_template(template_path)
    if suffix == ".docx":
        return _read_docx_template(template_path)
    raise TemplateReaderError("Selecione um modelo no formato .txt ou .docx.")


def _read_txt_template(path: Path) -> str:
    try:
        text = path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        text = path.read_text(encoding="cp1252")
    text = text.strip()
    if not text:
        raise TemplateReaderError("O arquivo de modelo está vazio.")
    return text


def _read_docx_template(path: Path) -> str:
    try:
        with ZipFile(path) as archive:
            document_xml = archive.read("word/document.xml")
    except (BadZipFile, KeyError, OSError) as exc:
        raise TemplateReaderError("Não foi possível abrir o modelo Word selecionado.") from exc

    root = ET.fromstring(document_xml)
    paragraphs: list[str] = []
    for paragraph in root.findall(".//w:body/w:p", WORD_NS):
        rendered = _render_paragraph(paragraph).strip()
        if rendered:
            paragraphs.append(rendered)

    if not paragraphs:
        raise TemplateReaderError("O modelo Word não possui texto para o corpo do e-mail.")
    return _normalize_template_text("\n\n".join(paragraphs))


def _normalize_template_text(template: str) -> str:
    return (
        template.replace("**R$****{Valor fev26}****,00**", "**{Valor fev26}**")
        .replace("R${Valor fev26},00", "{Valor fev26}")
        .replace("R$ {Valor fev26},00", "{Valor fev26}")
    )


def _render_paragraph(paragraph: ET.Element) -> str:
    parts: list[str] = []
    field_name: str | None = None
    in_field_result = False
    result_style = RunStyle()
    for run in paragraph.findall("w:r", WORD_NS):
        field_char = run.find("w:fldChar", WORD_NS)
        if field_char is not None:
            field_type = field_char.attrib.get(f"{{{WORD_NS['w']}}}fldCharType")
            if field_type == "begin":
                field_name = None
                in_field_result = False
                result_style = RunStyle()
            elif field_type == "separate":
                in_field_result = field_name is not None
            elif field_type == "end":
                if field_name is not None:
                    parts.append(_apply_style(f"{{{field_name}}}", result_style))
                field_name = None
                in_field_result = False
                result_style = RunStyle()
            continue

        instruction = _merge_field_name(run)
        if instruction is not None:
            field_name = instruction
            continue

        text = _run_text(run)
        if not text:
            continue
        if in_field_result:
            result_style = result_style.merge(RunStyle.from_run(run))
            continue
        parts.append(_apply_style(text, RunStyle.from_run(run)))
    return "".join(parts)


def _merge_field_name(run: ET.Element) -> str | None:
    instruction_text = "".join(
        element.text or "" for element in run.findall("w:instrText", WORD_NS)
    ).strip()
    if not instruction_text.startswith("MERGEFIELD"):
        return None
    parts = instruction_text.split()
    if len(parts) < 2:
        return None
    raw_name = parts[1].strip('"')
    return MERGE_FIELD_PLACEHOLDERS.get(raw_name, raw_name.replace("_", " "))


def _run_text(run: ET.Element) -> str:
    pieces: list[str] = []
    for child in run:
        tag = _local_name(child.tag)
        if tag == "t":
            pieces.append(child.text or "")
        elif tag in {"br", "cr"}:
            pieces.append("\n")
        elif tag == "tab":
            pieces.append("\t")
    return "".join(pieces)


def _apply_style(text: str, style: RunStyle) -> str:
    if style.underline:
        text = f"__{text}__"
    if style.italic:
        text = f"*{text}*"
    if style.bold:
        text = f"**{text}**"
    return text


class RunStyle:
    def __init__(
        self, *, bold: bool = False, italic: bool = False, underline: bool = False
    ) -> None:
        self.bold = bold
        self.italic = italic
        self.underline = underline

    @classmethod
    def from_run(cls, run: ET.Element) -> RunStyle:
        run_properties = run.find("w:rPr", WORD_NS)
        if run_properties is None:
            return cls()
        return cls(
            bold=run_properties.find("w:b", WORD_NS) is not None,
            italic=run_properties.find("w:i", WORD_NS) is not None,
            underline=run_properties.find("w:u", WORD_NS) is not None,
        )

    def merge(self, other: RunStyle) -> RunStyle:
        return RunStyle(
            bold=self.bold or other.bold,
            italic=self.italic or other.italic,
            underline=self.underline or other.underline,
        )


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]
