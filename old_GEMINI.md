# Microsoft Word Extractor

## Purpose

This script provides a command-line tool to convert Microsoft Word documents (.docx, .doc) into plain text files (.txt). It addresses challenges with complex Word document formatting, including tables and form fields embedded within content containers, by utilizing the `docx2txt` library for robust text extraction.

## Usage

The script `extractor.py` can be executed from the command line with two main commands:

### 1. Convert a single Word document

This command converts a specified Word document to a plain text file. If an output path is not provided, the script will automatically generate one by appending `_converted.txt` to the original file's base name.

```bash
python extractor.py convert-file <input_path> [<output_path>]
```

-   `<input_path>`: The full path to the Word document (.docx or .doc) to be converted.
-   `[<output_path>]`: (Optional) The desired full path for the output plain text file. If omitted, the output will be saved in the same directory as the input file with `_converted.txt` appended.

**Example:**
```bash
python extractor.py convert-file "C:\Documents\my_report.docx"
```
This will create `C:\Documents\my_report_converted.txt`.

### 2. Convert all Word documents in a folder

This command iterates through all Word documents (.docx, .doc) within a specified folder and converts each one to a plain text file. For each converted file, the output text file will be named `originalfilename_converted.txt` and saved in the same directory as the original Word document.

```bash
python extractor.py convert-folder <folder_path>
```

-   `<folder_path>`: The full path to the directory containing the Word documents to be converted.

**Example:**
```bash
python extractor.py convert-folder "C:\Documents\Projects"
```
This will convert all Word documents found in the `C:\Documents\Projects` folder.

## Dependencies

The script relies on the following Python libraries:

-   `python-docx`: Used for general Word document handling (though `docx2txt` is primarily used for text extraction).
-   `docx2txt`: Used for robust text extraction, especially effective with form fields and complex content controls.

These dependencies are listed in the `requirements.txt` file and can be installed using pip:

```bash
pip install -r requirements.txt
```