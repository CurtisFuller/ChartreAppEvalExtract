# Charter School Evaluation Comments Compilation Tool - Specification

## Project Overview
Create a Python command-line script that extracts reviewer comments from multiple Microsoft Word charter school evaluation documents and compiles them into a single organized Markdown document. The script must intelligently parse Word document tables, filter out standard evaluation matrix boilerplate text, and organize comments by section with reviewer attribution.

## Input Requirements

### File Structure
- **Review Documents**: Multiple Word (.docx) files in a user-specified folder
- **Naming Convention**: `SchoolName_Eval_ReviewerName.docx`
  - Example: `PineappleCove_Eval_JohnSmith.docx`
- **Word Document Form Style**: 
- **Template Files**: Separate folder containing evaluation matrix template files (for boilerplate text identification)
  - State model template
  - FCI-branded organization template (may have additional text beyond state model)

### Document Structure
Each evaluation document contains:
- Sections numbered and titled (e.g., "Section 1. Mission Guiding Principles and Purpose")
- Within each section:
  - **Evaluation Criteria**: Bulleted list (to be excluded from output)
  - **Standards Rating Table**: "Meets the Standard", "Partially Meets the Standard", "Does Not Meet the Standard" (to be excluded)
  - **Comments Table** with structure:
    - Row 1: "Strengths" (column 1) | "Reference" (column 2)
    - Row 2: Blank row (where reviewers add strength comments)
    - Row 3: "Concerns and Additional Questions" (column 1) | "Reference" (column 2)
    - Row 4: Blank row (where reviewers add concern comments)

**Important Notes**:
- Table formatting may vary (especially from Google Docs exports)
- Header text may have minor variations
- Comments are in the blank rows below headers
- Reference column contains page numbers (may not align perfectly with comments)

## Application Types and Sections

The script must support three application types:

### Standard Application (22 sections)
1. Mission Guiding Principles and Purpose
2. Target Population and Student Body
3. Educational Program Design
4. Curriculum and Instructional Design
5. Student Performance
6. Exceptional Students
7. English Language Learners
8. School Culture and Discipline
9. Supplemental Programming
10. Governance
11. Management and Staffing
12. Human Resources and Employment
13. Professional Development
14. Student Recruitment and Enrollment
15. Parent and Community Involvement
16. Facilities
17. Transportation Service
18. Food Service
19. School Safety and Security
20. Budget
21. Financial Management and Oversight
22. Start-Up Plan

### Virtual Application (18 sections)
1. Mission Guiding Principles and Purpose
2. Target Population and Student Body
3. Educational Program Design
4. Curriculum and Instructional Design
5. Student Performance
6. Exceptional Students
7. English Language Learners
8. School Culture and Discipline
9. Supplemental Programming
10. Governance
11. Management and Staffing
12. Human Resources and Employment
13. Professional Development
14. Student Recruitment and Enrollment
15. Parent and Community Involvement
16. Budget
17. Financial Management and Oversight
18. Start-Up Plan

### High Performing System Replication (12 sections)
1. Replication Overview
2. Mission Guiding Principles and Purpose
3. Educational Program, Curriculum, and Instructional Design
4. Student Performance
5. Student Recruitment and Enrollment
6. Management and Staffing
7. Facilities
8. Transportation Service
9. Food Service
10. School Safety and Security
11. Budget
12. Financial Management and Oversight

**Note**: Additional addendum sections may occasionally appear (e.g., "Addendum: Education Service Providers")

## Processing Requirements

### Intelligent Text Extraction
- **DO NOT** rely on exact string matching for headers or structure
- Parse Word document tables dynamically to identify comment sections
- Handle variations in formatting from different document sources
- Identify "Strengths" and "Concerns/Questions" sections within tables regardless of minor text variations

### Boilerplate Filtering
- Load template files from a separate template folder
- Extract all standard evaluation matrix text from templates
- Exclude any text from reviewer documents that matches template boilerplate
- This includes:
  - Evaluation criteria language
  - Standard instructions
  - Rating scale descriptions
  - Any other non-comment text from templates

### Comment Extraction Rules
1. **Extract only reviewer-added comments** from the blank rows in tables
2. **Include page number references** when available, formatted as: `Comment text [p. 23]`
3. **Skip blank sections** - if no comments provided, omit that section/subsection
4. **Attribute comments to reviewers** - extract reviewer name from filename
5. **Remove exact duplicate comments** - if identical text appears multiple times, include only once
6. **Keep similar comments** - do not attempt fuzzy matching; only remove exact duplicates

### Section Organization
- Output sections using the exact names provided in the specification lists above
- Process whatever sections are found in the documents
- Match section names flexibly (handle minor variations like "Section 3:" vs "Section 3.")
- If addendum sections are found, include them at the end

## User Interface Requirements

### Command-Line Prompts
When script is run, prompt for:

1. **Path to folder containing review documents**
   - Validate folder exists and contains .docx files

2. **Path to folder containing evaluation matrix templates**
   - Validate folder exists and contains template files

3. **Application type selection**
   - Prompt: "Select application type: (1) Standard, (2) Virtual, (3) High Performing Replication"
   - Use this to know which section list to expect

4. **Output filename**
   - Suggest default: `{SchoolName}_CompiledReviews.md`
   - Allow user to override

### Progress Indicators
- Display each file being processed
- Show count of comments extracted per reviewer
- Display any errors or warnings (e.g., files that couldn't be parsed)

## Output Format

### Markdown Structure
```markdown
# Charter Application Review Comments Compilation
## School Name: [Extracted from filename]

## Section 1. [Section Name]
### Strengths
- [Reviewer Name]: Comment text here [p. 23]
- [Reviewer Name]: Another comment [p. 45]

### Concerns
- [Reviewer Name]: Concern text here [p. 67]
- [Reviewer Name]: Another concern [p. 89]

## Section 2. [Section Name]
### Strengths
- [Reviewer Name]: Comment text here

### Concerns
- [Reviewer Name]: Concern text here

[Continue for all sections...]
```

### Output Rules
- Save to the same folder as the input review documents
- If a section has no comments from any reviewer, omit that section entirely
- If a section has Strengths but no Concerns (or vice versa), only include the populated subsection
- Preserve all unique comments with reviewer attribution

## Technical Implementation Notes

### Required Python Libraries
- **python-docx**: For parsing Word documents
- **pathlib**: For file path handling
- Standard library modules as needed

### Error Handling
- Gracefully handle corrupted or malformed Word documents
- Warn if template files cannot be loaded
- Continue processing if individual files fail
- Provide summary of any files that couldn't be processed

### Code Organization Suggestions
- Create functions for:
  - Loading and parsing templates (extract boilerplate text)
  - Parsing individual evaluation documents
  - Extracting reviewer name from filename
  - Identifying and extracting table contents
  - Filtering boilerplate text
  - Organizing comments by section and type
  - Generating Markdown output
- Use data structures (dictionaries/objects) to organize comments by section and type before output

### Performance Considerations
- Process template files once at startup
- Build boilerplate text index for efficient comparison
- Process review documents sequentially (5-8 files is manageable)

## Success Criteria
The script successfully:
1. Prompts for all necessary inputs with clear instructions
2. Loads and parses template files to identify boilerplate text
3. Extracts reviewer names from filenames correctly
4. Parses Word document tables regardless of formatting variations
5. Excludes all standard evaluation matrix language
6. Captures all unique reviewer comments with attribution
7. Includes page references when available
8. Organizes output by section using specified section names
9. Generates clean, readable Markdown output
10. Handles errors gracefully and reports any issues

## Example Usage Flow
```
$ python charter_review_compiler.py

Charter School Evaluation Comments Compiler
===========================================

Enter path to folder containing review documents: /path/to/reviews
Found 6 .docx files

Enter path to folder containing evaluation matrix templates: /path/to/templates
Loaded 2 template files

Select application type:
  1. Standard Application (22 sections)
  2. Virtual Application (18 sections)
  3. High Performing Replication (12 sections)
Enter selection (1-3): 1

Enter output filename [PineappleCove_CompiledReviews.md]: 

Processing reviews...
  ✓ PineappleCove_Eval_JohnSmith.docx (45 comments extracted)
  ✓ PineappleCove_Eval_JaneDoe.docx (38 comments extracted)
  ✓ PineappleCove_Eval_BobJones.docx (52 comments extracted)
  ✓ PineappleCove_Eval_SarahWilliams.docx (41 comments extracted)
  ✓ PineappleCove_Eval_MikeGarcia.docx (47 comments extracted)
  ✓ PineappleCove_Eval_LisaBrown.docx (39 comments extracted)

Removed 156 boilerplate text instances
Removed 8 duplicate comments
Compiled 254 unique comments across 22 sections

Output saved to: /path/to/reviews/PineappleCove_CompiledReviews.md
```

## Additional Context for Implementation

**Purpose**: This tool streamlines charter school application review processes at Florida Charter Institute. Currently, reviewers manually compile 5-8 evaluation documents per application to aggregate reviewer feedback. This tool automates that process, saving significant time during capacity interview preparation.

**User Skill Level**: The intended user is comfortable with command-line tools and has strong technical skills (Python, PHP, database work), so technical error messages and command-line interface are appropriate.

**Critical Success Factor**: The intelligent parsing of table structures is essential since document formatting will vary, especially when exported from Google Docs. Robust table detection and flexible text matching are more important than perfect efficiency.

## Future Enhancements (Optional)
- GUI interface for users less comfortable with command-line
- Batch processing of multiple schools
- Export to Word or PDF in addition to Markdown
- Statistical summary of comments per section
- Sentiment analysis or keyword extraction from comments
- Configuration file for custom section lists
