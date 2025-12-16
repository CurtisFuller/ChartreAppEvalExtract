from pathlib import Path
from tempfile import TemporaryDirectory
import sys
import zipfile

# Ensure project root on path for direct test execution
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from extractor import extract_docx_text

CONTENT_TYPES = """<?xml version='1.0' encoding='UTF-8'?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

RELS = """<?xml version='1.0' encoding='UTF-8'?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOCUMENT_TEMPLATE = """<?xml version='1.0' encoding='UTF-8'?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">
  <w:body>
    {content}
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>
"""


def _write_docx(path: Path, document_xml: str) -> None:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')


def test_extracts_checkbox_content_controls_from_table_cells():
    table_xml = """
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p>
            <w:sdt>
              <w:sdtPr>
                <w14:checkbox>
                  <w14:checked w14:val="1"/>
                </w14:checkbox>
              </w:sdtPr>
              <w:sdtContent>
                <w:p><w:r><w:t>Yes</w:t></w:r></w:p>
              </w:sdtContent>
            </w:sdt>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>Regular text</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    """
    document_xml = DOCUMENT_TEMPLATE.format(content=table_xml)

    with TemporaryDirectory() as tmp:
        path = Path(tmp) / "checkbox.docx"
        _write_docx(path, document_xml)
        text = extract_docx_text(path)

    assert "☒" in text
    assert "Yes" in text
    assert "Regular text" in text


def test_extracts_legacy_form_field_values_inside_tables():
    table_xml = """
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r>
              <w:ffData>
                <w:textInput>
                  <w:default w:val="LegacyValue"/>
                </w:textInput>
              </w:ffData>
            </w:r>
            <w:r><w:t>LegacyValue</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    """
    document_xml = DOCUMENT_TEMPLATE.format(content=table_xml)

    with TemporaryDirectory() as tmp:
        path = Path(tmp) / "legacy.docx"
        _write_docx(path, document_xml)
        text = extract_docx_text(path)

    assert "LegacyValue" in text


def test_form_values_keep_position_within_cell():
    table_xml = """
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r><w:t>Label:</w:t></w:r>
            <w:sdt>
              <w:sdtPr><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr>
              <w:sdtContent><w:p><w:r><w:t>Off</w:t></w:r></w:p></w:sdtContent>
            </w:sdt>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r><w:t>Name</w:t></w:r>
            <w:r><w:ffData><w:textInput><w:result w:val="Alice"/></w:textInput></w:ffData></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    """
    document_xml = DOCUMENT_TEMPLATE.format(content=table_xml)

    with TemporaryDirectory() as tmp:
        path = Path(tmp) / "order.docx"
        _write_docx(path, document_xml)
        text = extract_docx_text(path)

    lines = text.split("\n")
    assert lines[0].startswith("Label:")
    assert "☐" in lines[0]
    assert "Off" in lines[0]
    assert "Name" in lines[0]
    assert "Alice" in lines[0]


def test_legacy_checkbox_within_table_includes_state():
    table_xml = """
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r>
              <w:ffData>
                <w:checkBox>
                  <w:checked w:val="0"/>
                </w:checkBox>
              </w:ffData>
            </w:r>
            <w:r><w:t>Unchecked legacy box</w:t></w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r>
              <w:ffData>
                <w:checkBox>
                  <w:checked w:val="1"/>
                </w:checkBox>
              </w:ffData>
            </w:r>
            <w:r><w:t>Checked legacy box</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    """
    document_xml = DOCUMENT_TEMPLATE.format(content=table_xml)

    with TemporaryDirectory() as tmp:
        path = Path(tmp) / "legacy_checkbox.docx"
        _write_docx(path, document_xml)
        text = extract_docx_text(path)

    assert "☐" in text
    assert "☒" in text
    assert "Unchecked legacy box" in text
    assert "Checked legacy box" in text
