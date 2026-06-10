from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile


class TemplateReaderError(ValueError):
    """Friendly template reader error."""


WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


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
    return "\n\n".join(paragraphs)


def _render_paragraph(paragraph: ET.Element) -> str:
    parts: list[str] = []
    for run in paragraph.findall("w:r", WORD_NS):
        text = _run_text(run)
        if not text:
            continue
        parts.append(_apply_run_style(text, run))
    return "".join(parts)


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


def _apply_run_style(text: str, run: ET.Element) -> str:
    run_properties = run.find("w:rPr", WORD_NS)
    if run_properties is None:
        return text

    if run_properties.find("w:u", WORD_NS) is not None:
        text = f"__{text}__"
    if run_properties.find("w:i", WORD_NS) is not None:
        text = f"*{text}*"
    if run_properties.find("w:b", WORD_NS) is not None:
        text = f"**{text}**"
    return text


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]
