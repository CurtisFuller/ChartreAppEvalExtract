import os
import docx2txt

def convert_single_document(input_path=None, output_path=None):
    """
    Converts a single Word document to a plain text file.

    If no input_path is provided, the user is prompted to enter one.
    If no output_path is provided, it's generated from the input_path.
    """

    if not os.path.exists(input_path):
        print(f"Error: Input file not found at {input_path}")
        return

    if not output_path:
        base_name, _ = os.path.splitext(input_path)
        output_path = base_name + "_converted.txt"

    try:
        text = docx2txt.process(input_path)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        
        print(f"Successfully converted {input_path} to {output_path}")

    except Exception as e:
        print(f"An error occurred: {e}")

def convert_folder(folder_path):
    """
    Converts all Word documents in a folder to plain text files.
    """
    if not os.path.isdir(folder_path):
        print(f"Error: Folder not found at {folder_path}")
        return

    for filename in os.listdir(folder_path):
        if filename.lower().endswith((".doc", ".docx")):
            input_path = os.path.join(folder_path, filename)
            convert_single_document(input_path)

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python extractor.py <command> [<args>]")
        print("Commands:")
        print("  convert-file <input_path> [<output_path>]")
        print("  convert-folder <folder_path>")
        sys.exit(1)

    command = sys.argv[1]

    if command == "convert-file":
        if len(sys.argv) < 3:
            print("Usage: python extractor.py convert-file <input_path> [<output_path>]")
            sys.exit(1)
        input_file = sys.argv[2]
        output_file = sys.argv[3] if len(sys.argv) > 3 else None
        convert_single_document(input_file, output_file)
    elif command == "convert-folder":
        if len(sys.argv) < 3:
            print("Usage: python extractor.py convert-folder <folder_path>")
            sys.exit(1)
        folder = sys.argv[2]
        convert_folder(folder)
    else:
        print(f"Unknown command: {command}")
        sys.exit(1)
