#!/usr/bin/env python3
"""
Converts large Word.js API documentation Markdown files to JSON using local parsing.

This script is specifically designed for large, complex documentation files that
may be too large or complex for direct API-based conversion. It uses the md_parser
module to analyze the document structure and extract information hierarchically.

Usage:
    python jsonfy_large_docs.py --files Word.Body.md Word.Range.md Word.Paragraph.md Word.Document.md
    python jsonfy_large_docs.py --all-large  # Process all files > 50KB
    python jsonfy_large_docs.py --file Word.Body.md --output custom_output.json
"""

import argparse
import json
import sys
from pathlib import Path
from typing import List, Optional, Sequence

try:
    from md_parser import APIDocParser
except ModuleNotFoundError as exc:
    raise SystemExit(
        "Missing dependency: md_parser.py must be in the same directory."
    ) from exc


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert large Word JS class docs to JSON using local parsing."
    )

    parser.add_argument(
        "--docs-dir",
        default="api_docs",
        help="Directory containing Markdown files (default: api_docs)",
    )

    parser.add_argument(
        "--output-dir",
        default="jsonfied",
        help="Directory to write JSON files (default: jsonfied)",
    )

    # Input file selection
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "--files",
        nargs="+",
        help="Specific files to process (e.g., Word.Body.md Word.Range.md)",
    )
    group.add_argument(
        "--file",
        help="Single file to process",
    )
    group.add_argument(
        "--all-large",
        action="store_true",
        help="Process all files larger than --size-threshold",
    )

    parser.add_argument(
        "--size-threshold",
        type=int,
        default=50000,
        help="Size threshold in bytes for --all-large (default: 50000)",
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing JSON files",
    )

    parser.add_argument(
        "--output",
        help="Custom output file (only works with --file)",
    )

    parser.add_argument(
        "--pretty",
        action="store_true",
        default=True,
        help="Pretty-print JSON output (default: True)",
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be done without writing files",
    )

    return parser.parse_args(argv)


def find_large_files(docs_dir: Path, size_threshold: int) -> List[Path]:
    """Find all markdown files larger than the threshold."""
    large_files = []

    if not docs_dir.is_dir():
        return large_files

    for md_file in docs_dir.glob("*.md"):
        if md_file.stat().st_size > size_threshold:
            large_files.append(md_file)

    return sorted(large_files, key=lambda f: f.stat().st_size, reverse=True)


def process_file(
    input_path: Path,
    output_path: Path,
    overwrite: bool,
    pretty: bool,
    dry_run: bool
) -> bool:
    """Process a single documentation file."""

    # Check if output exists
    if output_path.exists() and not overwrite:
        print(f"[SKIP] {output_path.name} (already exists)", file=sys.stderr)
        return True

    # Read markdown
    try:
        markdown_text = input_path.read_text(encoding="utf-8")
    except Exception as exc:
        print(f"[ERROR] Failed to read {input_path}: {exc}", file=sys.stderr)
        return False

    # Parse document
    print(f"[PARSE] Parsing {input_path.name}...", file=sys.stderr)
    try:
        parser = APIDocParser(markdown_text)
        json_data = parser.to_json_schema()
    except Exception as exc:
        print(f"[ERROR] Failed to parse {input_path.name}: {exc}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return False

    if dry_run:
        print(f"[DRY] Would write to {output_path}", file=sys.stderr)
        # Print sample of parsed data
        print(f"  Class: {json_data.get('class', {}).get('name', 'Unknown')}")
        print(f"  Properties: {len(json_data.get('properties', []))}")
        print(f"  Methods: {len(json_data.get('methods', []))}")
        return True

    # Write JSON
    try:
        if pretty:
            json_text = json.dumps(json_data, ensure_ascii=False, indent=2)
        else:
            json_text = json.dumps(json_data, ensure_ascii=False)

        output_path.write_text(json_text, encoding="utf-8")
        print(f"[OK] {input_path.name} -> {output_path.name}", file=sys.stderr)

        # Print statistics
        print(f"  Class: {json_data.get('class', {}).get('name', 'Unknown')}")
        print(f"  Properties: {len(json_data.get('properties', []))}")
        print(f"  Methods: {len(json_data.get('methods', []))}")

        return True

    except Exception as exc:
        print(f"[ERROR] Failed to write {output_path}: {exc}", file=sys.stderr)
        return False


def main(argv: Sequence[str] | None = None) -> None:
    args = parse_args(argv)

    docs_dir = Path(args.docs_dir)
    output_dir = Path(args.output_dir)

    # Validate directories
    if not docs_dir.is_dir():
        raise SystemExit(f"Docs directory not found: {docs_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)

    # Determine which files to process
    files_to_process: List[Path] = []

    if args.file:
        # Single file mode
        file_path = docs_dir / args.file if not Path(args.file).is_absolute() else Path(args.file)
        if not file_path.exists():
            raise SystemExit(f"File not found: {file_path}")
        files_to_process = [file_path]

    elif args.files:
        # Multiple files mode
        for filename in args.files:
            file_path = docs_dir / filename if not Path(filename).is_absolute() else Path(filename)
            if not file_path.exists():
                print(f"[WARN] File not found: {file_path}", file=sys.stderr)
                continue
            files_to_process.append(file_path)

    elif args.all_large:
        # Find all large files
        files_to_process = find_large_files(docs_dir, args.size_threshold)
        if not files_to_process:
            print(f"No files larger than {args.size_threshold} bytes found in {docs_dir}")
            return

        print(f"Found {len(files_to_process)} large files:", file=sys.stderr)
        for f in files_to_process:
            size_kb = f.stat().st_size / 1024
            print(f"  - {f.name} ({size_kb:.1f} KB)", file=sys.stderr)

    if not files_to_process:
        raise SystemExit("No files to process")

    # Process files
    successes = 0
    failures = []

    for i, input_path in enumerate(files_to_process, start=1):
        print(f"\n[{i}/{len(files_to_process)}] Processing {input_path.name}...", file=sys.stderr)

        # Determine output path
        if args.output and len(files_to_process) == 1:
            output_path = Path(args.output)
        else:
            # Extract class name from filename (e.g., Word.Body.md -> Word.Body.json)
            output_filename = input_path.stem + ".json"
            output_path = output_dir / output_filename

        success = process_file(
            input_path,
            output_path,
            args.overwrite,
            args.pretty,
            args.dry_run
        )

        if success:
            successes += 1
        else:
            failures.append(input_path.name)

    # Summary
    print(f"\n{'='*60}", file=sys.stderr)
    if args.dry_run:
        print(f"Dry-run complete. Would process {len(files_to_process)} files.", file=sys.stderr)
    else:
        print(f"Completed: {successes} successes, {len(failures)} failures", file=sys.stderr)

        if failures:
            print(f"\nFailed files:", file=sys.stderr)
            for name in failures:
                print(f"  - {name}", file=sys.stderr)
            raise SystemExit(1)

        print(f"\nAll outputs saved to {output_dir}", file=sys.stderr)


if __name__ == "__main__":
    main()
