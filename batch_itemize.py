import os
import sys
from itemize import json_to_markdown

def batch_convert(input_dir='processed', output_dir='markdown_output'):
    """Batch convert all JSON files in a directory to Markdown"""

    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created output directory: {output_dir}")

    # Get all JSON files
    json_files = [f for f in os.listdir(input_dir) if f.endswith('.json')]

    if not json_files:
        print(f"No JSON files found in {input_dir}")
        return

    print(f"Found {len(json_files)} JSON files to convert")
    print("=" * 80)

    success_count = 0
    error_count = 0

    for filename in json_files:
        json_path = os.path.join(input_dir, filename)
        md_filename = os.path.splitext(filename)[0] + '.md'
        md_path = os.path.join(output_dir, md_filename)

        try:
            json_to_markdown(json_path, md_path)
            success_count += 1
        except Exception as e:
            print(f"Error converting {filename}: {e}")
            error_count += 1

    print("=" * 80)
    print(f"\nConversion complete!")
    print(f"  Successful: {success_count}")
    print(f"  Errors: {error_count}")
    print(f"  Total: {len(json_files)}")
    print(f"\nMarkdown files saved in: {output_dir}")

def main():
    if len(sys.argv) > 1:
        input_dir = sys.argv[1]
    else:
        input_dir = 'processed'

    if len(sys.argv) > 2:
        output_dir = sys.argv[2]
    else:
        output_dir = 'markdown_output'

    if not os.path.exists(input_dir):
        print(f"Error: Input directory {input_dir} does not exist")
        sys.exit(1)

    batch_convert(input_dir, output_dir)

if __name__ == '__main__':
    main()
