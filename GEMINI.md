# Charter Evaluation Compiler - Project Summary

## 1. Purpose
This project automates the extraction and compilation of reviewer comments from charter school application evaluation documents (`.docx`). The goal is to produce a single, aggregated Markdown report that groups "Strengths" and "Concerns" by application section, attributing each comment to the correct reviewer.

## 2. File Manifest
- **`charter_eval_compiler.py`**: The main, executable Python script that orchestrates the entire process.
- **`section_definitions.py`**: A crucial helper module containing dictionaries of all known section header variations for the three different application types (Standard, Virtual, High-Performing).
- **`extractor.py`**: The original, simpler script for text extraction. (Superseded by the main script but kept for history).
- **`requirements.txt`**: Lists the necessary Python packages (`python-docx`, `docx2txt`).
- **`evaluations/`**: A directory where input `.docx` evaluation files should be placed.
- **`templates/`**: A directory where the blank `.docx` template files should be placed. The script uses these to learn what "boilerplate" text to ignore.
- **`output/`**: The directory where the final `charter_evaluation_compilation.md` report is generated.
- **`debug/`**: A directory where the script saves intermediate `.txt` files. Each file shows the raw text extracted from a `.docx` document *after* cleaning, which is essential for debugging parsing issues.

## 3. How to Use
1. Place evaluation `.docx` files into the `evaluations/` directory.
2. Place the blank template `.docx` files into the `templates/` directory.
3. Run the script from the terminal:
   ```bash
   python charter_eval_compiler.py
   ```
4. The final report will be created at `output/charter_evaluation_compilation.md`.

## 4. Key Logic and Implementation Details

This section documents the final, working logic of the script, which was arrived at after several iterations of debugging.

### Text Extraction
- The script uses the **`docx2txt`** library to perform the initial conversion from `.docx` to plain text. This was chosen over `python-docx` for its superior ability to handle complex formatting and extract text from form fields.

### Boilerplate Removal & Section Header Preservation
- **Initial Problem:** The script was incorrectly removing section headers (e.g., "1. Mission...") because they were part of the boilerplate templates. This caused the script to find "0 sections" in every document.
- **Solution:** The final logic prioritizes headers. The script iterates through every line of the cleaned text from a document.
    1.  **First, it checks if the line is a known section header** by looking it up in a "master map" of all possible header variations. If it's a header, the line is kept, and a new section is started.
    2.  **Only if the line is *not* a header**, it then checks if the line is in the set of boilerplate text. If it is, the line is discarded.
    3.  This `header-first` approach ensures headers are never accidentally removed.

### Section Header Definitions
- **The Key Discovery:** The section headers present in the actual documents **do not contain the word "Section"** (e.g., they are `1. Mission...`, not `Section 1: Mission...`).
- The `section_definitions.py` file was updated with these "number-only" variations. This file contains three functions (`basic_model_app`, `virtual_model_app`, `high_performing_app`) that define all known header text for each canonical section name.
- A helper function, `_build_header_map()`, consumes these definitions and creates a single lookup dictionary that maps every possible variation to its canonical name (e.g., it maps both `"Section 1: Mission..."` and `"1. Mission..."` to the key `"01-Mission"`).

### Comment Parsing (`parse_section` function)
- After a document is split into sections, each section's text is passed to the `parse_section` function.
- This function uses a **state machine** to robustly parse comments. It handles two main formats:
    1.  **Style A:** Looks for "Strengths" and "Concerns and Additional Questions" headers and extracts the text blocks underneath them.
    2.  **Style B/C:** If Style A is not found, it goes line-by-line. When it finds a line that is just the keyword `Strength` or `Concern` (or `Question`, etc.), it enters a "mode" and gathers all subsequent lines as a single multi-line comment until it encounters the next keyword. This was critical to fixing the `IndexError` that occurred when the script expected comments to be on the same line as the keyword.
