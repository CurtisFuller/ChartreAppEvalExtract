#!/usr/bin/env python3
"""
Charter School Evaluation Comments Compiler - XML Version
Extracts reviewer comments from Word documents using raw XML parsing for form fields.
"""

import os
import sys
from pathlib import Path
from collections import defaultdict
from typing import Dict, List, Set, Tuple
import zipfile
import xml.etree.ElementTree as ET


# Namespace mappings for Office Open XML
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}

# Section lists (same as original)
SECTION_LISTS = {
    '1': {  # Standard Application
        'name': 'Standard Application',
        'sections': [
            'Mission Guiding Principles and Purpose',
            'Target Population and Student Body',
            'Educational Program Design',
            'Curriculum and Instructional Design',
            'Student Performance',
            'Exceptional Students',
            'English Language Learners',
            'School Culture and Discipline',
            'Supplemental Programming',
            'Governance',
            'Management and Staffing',
            'Human Resources and Employment',
            'Professional Development',
            'Student Recruitment and Enrollment',
            'Parent and Community Involvement',
            'Facilities',
            'Transportation Service',
            'Food Service',
            'School Safety and Security',
            'Budget',
            'Financial Management and Oversight',
            'Start-Up Plan',
        ]
    },
    '2': {  # Virtual Application
        'name': 'Virtual Application',
        'sections': [
            'Mission, Guiding Principles and Purpose',
            'Target Population and Student Body',
            'Educational Program Design',
            'Curriculum Plan',
            'Student Performance, Assessment and Evaluation',
            'Exceptional Students',
            'English Language Learners',
            'School Culture and Discipline',
            'Supplemental Programming',
            'Governance',
            'Management and Staffing',
            'Human Resources and Employment',
            'Professional Development',
            'Student Recruitment and Enrollment',
            'Parent and Community Involvement',
            'Budget',
            'Financial Management and Oversight',
            'Start-Up Plan',
        ]
    },
    '3': {  # High Performing System Replication
        'name': 'High Performing System Replication',
        'sections': [
            'Replication Overview',
            'Mission Guiding Principles and Purpose',
            'Educational Program, Curriculum, and Instructional Design',
            'Student Performance',
            'Student Recruitment and Enrollment',
            'Management and Staffing',
            'Facilities',
            'Transportation Service',
            'Food Service',
            'School Safety and Security',
            'Budget',
            'Financial Management and Oversight',
        ]
    }
}


def extract_text_from_xml_element(element):
    """Extract all text from an XML element, including nested text nodes."""
    texts = []
    # Get direct text
    if element.text:
        texts.append(element.text)
    # Get text from all children
    for child in element.iter():
        if child.text and child != element:
            texts.append(child.text)
        if child.tail:
            texts.append(child.tail)
    return ''.join(texts).strip()


def extract_table_data_from_xml(doc_xml_path):
    """Extract table data from Word document XML, including content controls."""
    try:
        tree = ET.parse(doc_xml_path)
        root = tree.getroot()

        tables_data = []

        # Find all tables
        for table in root.findall('.//w:tbl', NAMESPACES):
            table_rows = []

            # Find all rows in the table
            for row in table.findall('.//w:tr', NAMESPACES):
                row_cells = []

                # Find all cells in the row
                for cell in row.findall('.//w:tc', NAMESPACES):
                    # Extract text from the cell, including content controls (sdt)
                    cell_text_parts = []

                    # Look for regular paragraphs
                    for para in cell.findall('.//w:p', NAMESPACES):
                        para_text = extract_text_from_xml_element(para)
                        if para_text:
                            cell_text_parts.append(para_text)

                    # Look for content controls (sdt elements) which may contain form fields
                    for sdt in cell.findall('.//w:sdt', NAMESPACES):
                        sdt_text = extract_text_from_xml_element(sdt)
                        if sdt_text and sdt_text not in cell_text_parts:
                            cell_text_parts.append(sdt_text)

                    cell_text = '\n'.join(cell_text_parts).strip()
                    row_cells.append(cell_text)

                if row_cells:  # Only add non-empty rows
                    table_rows.append(row_cells)

            if table_rows:
                tables_data.append(table_rows)

        return tables_data
    except Exception as e:
        print(f"Error parsing XML: {e}")
        return []


def process_docx_with_xml(docx_path):
    """Process a .docx file by extracting its XML and parsing tables."""
    try:
        # .docx files are zip archives
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            # Extract document.xml to temp location
            import tempfile
            with tempfile.TemporaryDirectory() as temp_dir:
                # Extract the main document XML
                xml_content = zip_ref.read('word/document.xml')
                xml_path = Path(temp_dir) / 'document.xml'
                xml_path.write_bytes(xml_content)

                # Parse the XML
                return extract_table_data_from_xml(xml_path)
    except Exception as e:
        print(f"Error processing DOCX: {e}")
        return []


def extract_reviewer_name(filename: str) -> str:
    """Extract reviewer name from filename."""
    parts = filename.replace('.docx', '').split('_')
    if len(parts) >= 3:
        # Handle both "Eval" and "Evanl" (typo in filenames)
        for i, part in enumerate(parts):
            if part.lower() in ['eval', 'evanl']:
                if i + 1 < len(parts):
                    return parts[i + 1]
    return 'Unknown Reviewer'


def extract_school_name(filename: str) -> str:
    """Extract school name from filename."""
    parts = filename.replace('.docx', '').split('_')
    if len(parts) >= 3:
        for i, part in enumerate(parts):
            if part.lower() in ['eval', 'evanl']:
                return '_'.join(parts[:i])
    return 'Unknown School'


def main():
    """Main program execution."""
    print("Charter School Evaluation Comments Compiler (XML Version)")
    print("=" * 61)
    print()

    # Prompt for review documents folder
    review_folder_path = input("Enter path to folder containing review documents: ").strip()
    review_folder = Path(review_folder_path)

    if not review_folder.exists() or not review_folder.is_dir():
        print(f"Error: Folder not found: {review_folder_path}")
        sys.exit(1)

    docx_files = [f for f in review_folder.glob('*.docx') if not f.name.startswith('~$')]

    if not docx_files:
        print(f"Error: No .docx files found in {review_folder_path}")
        sys.exit(1)

    print(f"Found {len(docx_files)} .docx file(s)")
    print()

    # Process first file as a test
    print("Testing XML extraction on first file...")
    test_file = docx_files[0]
    print(f"Processing: {test_file.name}")
    print()

    tables = process_docx_with_xml(test_file)

    print(f"Found {len(tables)} tables")
    print()

    # Display first few tables
    for idx, table in enumerate(tables[:5]):
        print(f"Table {idx}: {len(table)} rows x {len(table[0]) if table else 0} columns")
        for row_idx, row in enumerate(table[:3]):  # Show first 3 rows
            print(f"  Row {row_idx}: {row}")
        if len(table) > 3:
            print(f"  ... and {len(table) - 3} more rows")
        print()


if __name__ == '__main__':
    main()
