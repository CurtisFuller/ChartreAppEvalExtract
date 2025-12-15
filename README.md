# Charter School Evaluation Comments Compiler

A Python command-line tool that extracts reviewer comments from multiple charter school evaluation documents and compiles them into a single organized Markdown document.

## Features

- **Intelligent Document Parsing**: Automatically extracts comments from evaluation documents
- **Boilerplate Filtering**: Removes standard evaluation matrix text using template files
- **Reviewer Attribution**: Tracks and attributes comments to individual reviewers
- **Duplicate Removal**: Eliminates exact duplicate comments across reviewers
- **Flexible Section Support**: Handles three application types (Standard, Virtual, High Performing Replication)
- **Progress Indicators**: Shows real-time processing status and statistics

## Available Versions

### 1. **Plain Text Version (RECOMMENDED)** - `charter_review_compiler_plaintext.py`
This is the most reliable method. It automatically extracts text from Word documents.

**How it works:**
- Accepts both .docx and .txt files
- Automatically converts .docx files to plain text internally
- No manual export needed!

**Advantages:**
- Works with all Word form field types (unlike the XML parser)
- More reliable than parsing complex Word structures
- Faster processing
- Can process .docx files directly (no manual conversion required)

### 2. **Word Document Version** - `charter_review_compiler.py`
Works directly with Word files (.docx) but may have issues with certain form field types.

### 3. **XML Parsing Version** - `charter_review_compiler_v2.py`
Experimental version that attempts to parse Word XML structures directly.

## Requirements

- Python 3.7 or higher
- For .docx versions: python-docx library

## Installation

1. Clone this repository or download the files
2. Install dependencies (only needed for .docx versions):

```bash
pip install -r requirements.txt
```

## File Naming Convention

Review documents should follow this pattern (the script will try to extract names even if the pattern isn't exact):
```
SchoolName_Eval_ReviewerName.docx  (or .txt)
```

Examples:
- `PineappleCove_Eval_JohnSmith.txt`
- `OaklandAcademy_Eval_JaneDoe.txt`

## Usage

### Plain Text Version (Recommended)

```bash
python charter_review_compiler_plaintext.py
```

### Word Document Version

```bash
python charter_review_compiler.py
```

The script will prompt you for:

1. **Path to review documents folder**: Directory containing the reviewer evaluation files
2. **Path to template folder**: Directory containing evaluation matrix template files
3. **Application type**: Choose from:
   - Standard Application (22 sections)
   - Virtual Application (18 sections)
   - High Performing System Replication (12 sections)
4. **Output filename**: Name for the compiled Markdown file (default suggested)

### Example Session

```
Charter School Evaluation Comments Compiler
===========================================

Enter path to folder containing review documents: /path/to/reviews
Found 6 .docx files

Enter path to folder containing evaluation matrix templates: /path/to/templates
Loaded 2 template files

Select application type:
  1. Standard Application (22 sections)
  2. Virtual Application (18 sections)
  3. High Performing System Replication (12 sections)
Enter selection (1-3): 1

Enter output filename [PineappleCove_CompiledReviews.md]:

Processing reviews...
  ✓ PineappleCove_Eval_JohnSmith.docx (45 comments extracted)
  ✓ PineappleCove_Eval_JaneDoe.docx (38 comments extracted)
  ✓ PineappleCove_Eval_BobJones.docx (52 comments extracted)

Removed 156 boilerplate text instances
Removed 8 duplicate comments
Compiled 254 unique comments across 22 sections

Output saved to: /path/to/reviews/PineappleCove_CompiledReviews.md
```

## Output Format

The compiled Markdown file organizes comments by section:

```markdown
# Charter Application Review Comments Compilation
## School Name: PineappleCove

## Section 1. Mission Guiding Principles and Purpose
### Strengths
- JohnSmith: Comment text here [p. 23]
- JaneDoe: Another comment [p. 45]

### Concerns
- BobJones: Concern text here [p. 67]

## Section 2. Target Population and Student Body
### Strengths
- JaneDoe: Comment text here

...
```

## Document Structure Requirements

The script expects evaluation documents with:

1. Numbered sections (e.g., "Section 1. Mission Guiding Principles and Purpose")
2. Comment tables with:
   - "Strengths" header row
   - Blank row for strength comments
   - "Concerns and Additional Questions" header row
   - Blank row for concern comments
   - Optional "Reference" column for page numbers

## Error Handling

The script will:
- Continue processing if individual files fail
- Display warnings for corrupted or malformed documents
- Provide a summary of any processing issues
- Skip sections with no comments

## Troubleshooting

**No comments extracted:**
- Verify document table structure matches expected format
- Check that template files are correctly identifying boilerplate text
- Ensure review documents follow the naming convention

**Missing sections:**
- Sections with no comments are omitted from output
- Verify section names match the expected list for the application type

**Import errors:**
- Ensure python-docx is installed: `pip install python-docx`
- Verify Python version is 3.7 or higher

## Development Notes

The script uses:
- `python-docx` for Word document parsing
- `pathlib` for cross-platform file handling
- Dynamic table parsing to handle formatting variations
- Set-based duplicate detection for efficiency

## License

This tool is developed for Florida Charter Institute's internal use.

## Support

For issues or questions, contact your technical administrator.
