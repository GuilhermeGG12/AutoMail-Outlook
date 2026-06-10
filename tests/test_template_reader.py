from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

from mailmerge_assistant.template_reader import read_email_body_template


def test_read_txt_template(tmp_path: Path) -> None:
    path = tmp_path / "modelo.txt"
    path.write_text("Olá **{Proprietário/Dirigente}**", encoding="utf-8")

    assert read_email_body_template(path) == "Olá **{Proprietário/Dirigente}**"


def test_read_docx_template_preserves_basic_run_styles(tmp_path: Path) -> None:
    path = tmp_path / "modelo.docx"
    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Olá </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>{Proprietário/Dirigente}</w:t></w:r>
      <w:r><w:t>, veja </w:t></w:r>
      <w:r><w:rPr><w:i/></w:rPr><w:t>{RAZÃO SOCIAL}</w:t></w:r>
      <w:r><w:t> e </w:t></w:r>
      <w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>{Valor fev26}</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""
    with ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", document_xml)

    template = read_email_body_template(path)

    assert template == (
        "Olá **{Proprietário/Dirigente}**, veja *{RAZÃO SOCIAL}* e __{Valor fev26}__"
    )
