import os
import re
import docx2txt
import glob
from section_definitions import basic_model_app, virtual_model_app, high_performing_app

def extract_text_from_docx(file_path):
    """Extracts raw text from a .docx file using docx2txt."""
    try:
        text = docx2txt.process(file_path)
        return text
    except Exception as e:
        print(f"Error extracting text from {file_path}: {e}")
        return None

def clean_text(text):
    """
    Normalizes whitespace by replacing tabs and multiple spaces with a single space.
    """
    if not text:
        return ""
    text = text.replace('\t', ' ')
    while '  ' in text:
        text = text.replace('  ', ' ')
    return text.strip()

def load_templates(template_dir):
    """
    Loads and cleans boilerplate text from template files in the given directory.
    Returns a set of unique, cleaned lines from all templates.
    """
    boilerplate_lines = set()
    if not os.path.isdir(template_dir):
        print(f"Warning: Template directory '{template_dir}' not found. No boilerplate will be removed.")
        return boilerplate_lines

    template_files = glob.glob(os.path.join(template_dir, '*.docx'))
    if not template_files:
        print(f"Warning: No .docx files found in template directory '{template_dir}'. No boilerplate will be removed.")
        return boilerplate_lines

    for template_file in template_files:
        raw_text = extract_text_from_docx(template_file)
        if raw_text:
            cleaned_text = clean_text(raw_text)
            for line in cleaned_text.splitlines():
                boilerplate_lines.add(line)
    
    print(f"Loaded {len(boilerplate_lines)} unique boilerplate lines from {len(template_files)} templates.")
    return boilerplate_lines

def _build_header_map():
    """Builds a master map from every header variation to its canonical name."""
    master_map = {}
    header_defs = {
        "standard": basic_model_app(),
        "virtual": virtual_model_app(),
        "high_performing": high_performing_app()
    }
    
    for app_type, sections in header_defs.items():
        for canonical_name, variations in sections.items():
            for variation in variations:
                master_map[clean_text(variation)] = canonical_name
    return master_map

def _detect_application_type(text):
    """
    Detects the application type based on the presence of specific keywords.
    """
    text_lower = text.lower()
    if "virtual application" in text_lower:
        return "virtual"
    if "high performing replication" in text_lower:
        return "high_performing"
    return "standard"

def _format_page_numbers(comment_text):
    """Finds and reformats page number mentions in comment text."""
    pattern = r'\b(?:page|p\. |p\.|pg\.)\s*(\d+)(?:\s*-\s*(\d+))?\b'
    def replace_match(match):
        start_page = match.group(1)
        end_page = match.group(2)
        if end_page:
            return f'[p. {start_page}-{end_page}]'
        return f'[p. {start_page}]'
    return re.sub(pattern, replace_match, comment_text, flags=re.IGNORECASE)

def parse_section(section_text, reviewer_name):
    """
    Parses a single section's text to extract strengths and concerns
    based on Style A or a more complex, multi-line Style B/C formatting.
    """
    strengths = []
    concerns = []

    # Attempt Style A: A block of strengths followed by a block of concerns.
    style_a_strength_pattern = re.compile(r'Strengths\n(.*?)(?=\nConcerns and Additional Questions|$)', re.DOTALL | re.IGNORECASE)
    style_a_concern_pattern = re.compile(r'Concerns and Additional Questions\n(.*)', re.DOTALL | re.IGNORECASE)

    match_strength_a = style_a_strength_pattern.search(section_text)
    match_concern_a = style_a_concern_pattern.search(section_text)

    if match_strength_a or match_concern_a:
        if match_strength_a:
            # Treat the entire block as one comment, then split by lines that seem to be new points if needed.
            # For now, let's assume a simple line split works.
            for line in match_strength_a.group(1).strip().splitlines():
                if line.strip():
                    strengths.append(line.strip())
        if match_concern_a:
            for line in match_concern_a.group(1).strip().splitlines():
                if line.strip():
                    concerns.append(line.strip())
    else:
        # Fallback to Style B/C: A state machine for line-by-line parsing.
        current_mode = None  # Can be 'strength' or 'concern'
        current_comment_lines = []

        def save_current_comment():
            if current_comment_lines:
                comment = "\n".join(current_comment_lines).strip()
                if comment:
                    if current_mode == 'strength':
                        strengths.append(comment)
                    elif current_mode == 'concern':
                        concerns.append(comment)
                current_comment_lines.clear()

        for line in section_text.splitlines():
            line = line.strip()
            if not line:
                continue
            
            line_lower = line.lower()

            if line_lower == 'strength':
                save_current_comment()
                current_mode = 'strength'
            elif line_lower == 'concern' or line_lower == 'question' or line_lower == 'follow up' or line_lower == 'improvement':
                save_current_comment()
                current_mode = 'concern'
            elif line_lower.startswith("strength:"):
                save_current_comment()
                current_mode = 'strength'
                comment_text = line.split(':', 1)[1].strip()
                if comment_text:
                    current_comment_lines.append(comment_text)
            elif line_lower.startswith("concern:") or line_lower.startswith("question:") or line_lower.startswith("follow up:") or line_lower.startswith("improvement:"):
                save_current_comment()
                current_mode = 'concern'
                comment_text = line.split(':', 1)[1].strip()
                if comment_text:
                    current_comment_lines.append(comment_text)
            elif current_mode:
                # This line is part of the ongoing comment
                current_comment_lines.append(line)

        # Save any remaining comment after the loop finishes
        save_current_comment()

    # Apply page number formatting and deduplicate
    unique_strengths = []
    seen_strengths_text = set()
    for s_text in strengths:
        formatted_s_text = _format_page_numbers(s_text)
        if formatted_s_text and formatted_s_text not in seen_strengths_text:
            unique_strengths.append({'reviewer': reviewer_name, 'comment': formatted_s_text})
            seen_strengths_text.add(formatted_s_text)

    unique_concerns = []
    seen_concerns_text = set()
    for c_text in concerns:
        formatted_c_text = _format_page_numbers(c_text)
        if formatted_c_text and formatted_c_text not in seen_concerns_text:
            unique_concerns.append({'reviewer': reviewer_name, 'comment': formatted_c_text})
            seen_concerns_text.add(formatted_c_text)

    return {'strengths': unique_strengths, 'concerns': unique_concerns}

def generate_markdown_report(data, output_path, canonical_order):
    """Generates the final Markdown report from the aggregated data."""
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("# Charter Application Review Comments Compilation\n\n")

        for section_key in canonical_order:
            if section_key in data:
                comments_in_section = data[section_key]
                if comments_in_section['strengths'] or comments_in_section['concerns']:
                    section_title = comments_in_section.get('title', section_key) 
                    f.write(f"## {section_title}\n\n")

                    if comments_in_section['strengths']:
                        f.write("### Strengths\n")
                        sorted_strengths = sorted(comments_in_section['strengths'], key=lambda x: x['reviewer'])
                        for comment_entry in sorted_strengths:
                            f.write(f"- {comment_entry['reviewer']}: {comment_entry['comment']}\n")
                        f.write("\n")

                    if comments_in_section['concerns']:
                        f.write("### Concerns\n")
                        sorted_concerns = sorted(comments_in_section['concerns'], key=lambda x: x['reviewer'])
                        for comment_entry in sorted_concerns:
                            f.write(f"- {comment_entry['reviewer']}: {comment_entry['comment']}\n")
                        f.write("\n")
    print(f"\nMarkdown report generated at: {output_path}")

def main():
    """
    Main function to orchestrate the compilation process.
    """
    input_dir = "evaluations"
    template_dir = "templates"
    output_dir = "output"
    debug_dir = "debug"
    
    for d in [output_dir, debug_dir]:
        if not os.path.exists(d):
            os.makedirs(d)

    print("--- Starting Charter Evaluation Compiler ---")

    # 1. Load Templates & Build Header Map
    boilerplate = load_templates(template_dir)
    header_map = _build_header_map()

    # 2. Process Files
    all_comments = {}
    
    evaluation_files = glob.glob(os.path.join(input_dir, '*.docx'))
    if not evaluation_files:
        print(f"No .docx files found in '{input_dir}'. Exiting.")
        return
        
    canonical_order = list(basic_model_app().keys())

    for file_path in evaluation_files:
        print(f"\nProcessing file: {os.path.basename(file_path)}")
        
        filename = os.path.basename(file_path)
        match = re.match(r'(.+)_Ev(al|anl)_(.+)\.docx', filename, re.IGNORECASE)
        if not match:
            print(f"Warning: Filename '{filename}' does not match expected pattern. Skipping.")
            continue
        
        school_name = match.group(1).replace('_', ' ')
        reviewer_name = match.group(3).replace('_', ' ')
        
        print(f"  School: {school_name}, Reviewer: {reviewer_name}")

        raw_text = extract_text_from_docx(file_path)
        if not raw_text:
            continue
        
        # Get all cleaned lines from the document
        all_lines = clean_text(raw_text).splitlines()
        
        # Save the fully cleaned text (before any filtering) to debug
        debug_filename = os.path.splitext(filename)[0] + '.txt'
        with open(os.path.join(debug_dir, debug_filename), 'w', encoding='utf-8') as f:
            f.write("\n".join(all_lines))
        print(f"  Saved cleaned text to debug/{debug_filename}")

        # New, more robust section splitting logic
        sections = {}
        current_section_key = None
        current_content = []

        for line in all_lines:
            # Check if the line is a header FIRST
            if line in header_map:
                # If we were in a section, save its content
                if current_section_key:
                    sections[current_section_key] = "\n".join(current_content)
                
                # Start the new section
                current_section_key = header_map[line]
                current_content = [] # Reset content for the new section
            
            # If it's not a header, check if it's NOT boilerplate, then add it
            elif current_section_key and line not in boilerplate:
                current_content.append(line)
        
        # Save the very last section's content after the loop finishes
        if current_section_key:
            sections[current_section_key] = "\n".join(current_content)

        print(f"  Found {len(sections)} sections in the document.")

        for section_key, section_content in sections.items():
            parsed_comments = parse_section(section_content, reviewer_name)
            
            if parsed_comments['strengths'] or parsed_comments['concerns']:
                print(f"    -> Parsed Section '{section_key}': {len(parsed_comments['strengths'])} strengths, {len(parsed_comments['concerns'])} concerns.")
            
            if section_key not in all_comments:
                # Get the full, pretty title for the report
                full_title = "Unknown Section"
                for app_type in [basic_model_app, virtual_model_app, high_performing_app]:
                    if section_key in app_type():
                        full_title = app_type()[section_key][0]
                        break
                all_comments[section_key] = {'title': full_title, 'strengths': [], 'concerns': []}

            for s_comment in parsed_comments['strengths']:
                if s_comment['comment'] not in [c['comment'] for c in all_comments[section_key]['strengths']]:
                    all_comments[section_key]['strengths'].append(s_comment)
            
            for c_comment in parsed_comments['concerns']:
                if c_comment['comment'] not in [c['comment'] for c in all_comments[section_key]['concerns']]:
                    all_comments[section_key]['concerns'].append(c_comment)

    # 3. Generate Report
    report_path = os.path.join(output_dir, "charter_evaluation_compilation.md")
    generate_markdown_report(all_comments, report_path, canonical_order)
    
    print("\n--- Script processing complete. ---")


if __name__ == "__main__":
    main()
