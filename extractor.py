from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"


def _iter_text(element):
    for node in element.iter():
        if node.tag.endswith("}t") and node.text:
            yield node.text


def _checkbox_state(element) -> str | None:
    for node in element.iter():
        if node.tag.endswith("}checked"):
            val = node.attrib.get(f"{{{W14_NS}}}val") or node.attrib.get(f"{{{W_NS}}}val")
            if val is None:
                return "☒"
            return "☒" if val in {"1", "true", "on", "True"} else "☐"
    return None


def _extract_form_field_values(cell_elem: ET.Element) -> list[str]:
    values: list[str] = []
    for elem in cell_elem.iter():
        if elem.tag.endswith("}sdt"):
            checkbox = _checkbox_state(elem)
            if checkbox:
                values.append(checkbox)
                continue
            text = "".join(_iter_text(elem)).strip()
            if text:
                values.append(text)
        elif elem.tag.endswith("}ffData"):
            for default in elem.iter():
                if default.tag.endswith("}default"):
                    val = default.attrib.get(f"{{{W_NS}}}val")
                    if val:
                        values.append(val)
    return values


def _cell_text(cell_elem: ET.Element) -> str:
    base_text = "".join(_iter_text(cell_elem)).strip()
    form_values = _extract_form_field_values(cell_elem)
    extras = [val for val in form_values if val and val not in base_text]
    parts = [base_text] if base_text else []
    parts.extend(extras)
    return " ".join(parts).strip()


def _parse_table(table_elem: ET.Element, lines: list[str]) -> None:
    for row in table_elem.findall(f".//{{{W_NS}}}tr"):
        cell_texts = []
        for cell in row.findall(f"{{{W_NS}}}tc"):
            cell_texts.append(_cell_text(cell))
        joined = "\t".join(cell_texts).strip()
        if joined:
            lines.append(joined)


def _parse_body(body: ET.Element, lines: list[str]) -> None:
    for child in body:
        tag = child.tag
        if tag.endswith("}p"):
            text = "".join(_iter_text(child)).strip()
            if text:
                lines.append(text)
        elif tag.endswith("}tbl"):
            _parse_table(child, lines)


def extract_docx_text(docx_path: Path | str, output_path: Path | str | None = None) -> str:
    """Extract text from a DOCX file including form-field values in tables."""
    docx_path = Path(docx_path)
    with zipfile.ZipFile(docx_path) as zf:
        with zf.open("word/document.xml") as doc_xml:
            tree = ET.parse(doc_xml)
    root = tree.getroot()
    body = root.find(f"{{{W_NS}}}body")
    lines: list[str] = []
    if body is not None:
        _parse_body(body, lines)

    result = "\n".join(lines)
    if output_path is not None:
        Path(output_path).write_text(result, encoding="utf-8")
    return result


__all__ = ["extract_docx_text"]


def _build_cli() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Extract text from a DOCX file (including form-field values in tables) "
            "and either print it to stdout or write it to an output file."
        )
    )
    parser.add_argument("docx_path", type=Path, help="Path to the DOCX file to extract")
    parser.add_argument(
        "output_path",
        type=Path,
        nargs="?",
        help="Optional path to write the extracted text. Prints to stdout if omitted.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _build_cli().parse_args(argv)
    try:
        text = extract_docx_text(args.docx_path, args.output_path)
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    if args.output_path is None:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
