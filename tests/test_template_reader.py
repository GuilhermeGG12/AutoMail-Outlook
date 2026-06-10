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
    _write_docx(path, document_xml)

    template = read_email_body_template(path)

    assert template == (
        "Olá **{Proprietário/Dirigente}**, veja *{RAZÃO SOCIAL}* e __{Valor fev26}__"
    )


def test_read_docx_template_uses_merge_fields_not_displayed_values(tmp_path: Path) -> None:
    path = tmp_path / "modelo.docx"
    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Olá </w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> MERGEFIELD ProprietárioDirigente </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Ana Flávia</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:r><w:t>, valor </w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> MERGEFIELD Valor_fev26 </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>R$ 300,00</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:body>
</w:document>"""
    _write_docx(path, document_xml)

    template = read_email_body_template(path)

    assert template == "Olá **{Proprietário/Dirigente}**, valor **{Valor fev26}**"
    assert "Ana Flávia" not in template
    assert "R$ 300,00" not in template


def test_read_docx_template_normalizes_split_currency_field(tmp_path: Path) -> None:
    path = tmp_path / "modelo.docx"
    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Valor (</w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>R$</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> MERGEFIELD Valor_fev26 </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>300</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>,00</w:t></w:r>
      <w:r><w:t>)</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""
    _write_docx(path, document_xml)

    template = read_email_body_template(path)

    assert template == "Valor (**{Valor fev26}**)"


def _write_docx(path: Path, document_xml: str) -> None:
    with ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", document_xml)
