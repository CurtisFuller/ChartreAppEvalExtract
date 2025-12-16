from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"


def _checkbox_state(element) -> str | None:
    for node in element.iter():
        if node.tag.endswith("}checked"):
            val = node.attrib.get(f"{{{W14_NS}}}val") or node.attrib.get(f"{{{W_NS}}}val")
            if val is None:
                return "☒"
            return "☒" if val in {"1", "true", "on", "True"} else "☐"
    return None


def _legacy_form_value(ff_data: ET.Element) -> str | None:
    def checkbox_symbol(val: str | None) -> str:
        if val is None:
            return "☐"
        return "☒" if val in {"1", "true", "on", "True"} else "☐"

    for node in ff_data.iter():
        if node.tag.endswith("}checkBox"):
            checked = None
            default = None
            for child in node:
                if child.tag.endswith("}checked"):
                    checked = child.attrib.get(f"{{{W_NS}}}val")
                elif child.tag.endswith("}default"):
                    default = child.attrib.get(f"{{{W_NS}}}val")
            return checkbox_symbol(checked if checked is not None else default)
        if node.tag.endswith("}result"):
            val = node.attrib.get(f"{{{W_NS}}}val")
            if val:
                return val
        if node.tag.endswith("}default"):
            val = node.attrib.get(f"{{{W_NS}}}val")
            if val:
                return val
    return None


def _gather_text(element: ET.Element) -> str:
    tag = element.tag
    if tag.endswith("}t"):
        return element.text or ""
    if tag.endswith("}tab"):
        return "\t"
    if tag.endswith("}br"):
        return "\n"
    if tag.endswith("}sdt"):
        checkbox = _checkbox_state(element)
        content = element.find(f"{{{W_NS}}}sdtContent")
        content_text = _gather_text(content) if content is not None else ""
        if checkbox:
            combined = f"{checkbox} {content_text}".strip()
            return combined
        return content_text
    if tag.endswith("}ffData"):
        val = _legacy_form_value(element)
        return val or ""

    parts: list[str] = []
    for child in element:
        text = _gather_text(child)
        if text:
            parts.append(text)
    return "".join(parts)


def _cell_text(cell_elem: ET.Element) -> str:
    parts: list[str] = []
    for child in cell_elem:
        if child.tag.endswith("}p"):
            para_text = _gather_text(child).strip()
            if para_text:
                parts.append(para_text)
        elif child.tag.endswith("}tbl"):
            nested_lines: list[str] = []
            _parse_table(child, nested_lines)
            if nested_lines:
                parts.append("\n".join(nested_lines))
    return "\n".join(parts).strip()


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
            text = _gather_text(child).strip()
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
