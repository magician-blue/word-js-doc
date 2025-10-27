#!/usr/bin/env python3
"""
Process JSON files to generate or enhance examples for properties and methods.

This script reads JSON files from the jsonfied/ directory and:
1. For properties/methods without examples: generates new examples (description + usage_code)
2. For examples with usage_code but no description: generates description
3. Saves enhanced JSON files to the processed/ directory

Usage:
    python process.py --file Word.Body.json
    python process.py --all
    python process.py --files Word.Body.json Word.Range.json
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence

try:
    from openai import OpenAI
except ModuleNotFoundError as exc:
    raise SystemExit(
        "Missing dependency: install the OpenAI SDK first, e.g. `pip install openai`."
    ) from exc


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Process JSON files to generate/enhance examples for properties and methods."
    )

    parser.add_argument(
        "--input-dir",
        default="jsonfied",
        help="Directory containing JSON files to process (default: jsonfied)",
    )

    parser.add_argument(
        "--output-dir",
        default="processed",
        help="Directory to save processed JSON files (default: processed)",
    )

    # Input file selection
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "--file",
        help="Process a single JSON file",
    )
    group.add_argument(
        "--files",
        nargs="+",
        help="Process multiple JSON files",
    )
    group.add_argument(
        "--all",
        action="store_true",
        help="Process all JSON files in input directory",
    )

    parser.add_argument(
        "--model",
        default="gpt-4o-mini",
        help="OpenAI model to use (default: gpt-4o-mini)",
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing processed files",
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be done without calling API or writing files",
    )

    parser.add_argument(
        "--skip-properties",
        action="store_true",
        help="Skip processing properties (only process methods)",
    )

    parser.add_argument(
        "--skip-methods",
        action="store_true",
        help="Skip processing methods (only process properties)",
    )

    parser.add_argument(
        "--max-per-file",
        type=int,
        help="Maximum number of examples to generate per file (for testing)",
    )

    return parser.parse_args(argv)


class ExampleGenerator:
    """Generate examples and descriptions using OpenAI API."""

    def __init__(self, client: OpenAI, model: str):
        self.client = client
        self.model = model

    def generate_property_example(
        self,
        class_name: str,
        class_description: str,
        property_info: Dict
    ) -> Dict:
        """Generate an example for a property."""
        prompt = self._build_property_example_prompt(
            class_name, class_description, property_info
        )

        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": "You are an expert at creating concise, practical API usage examples."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
        )

        content = response.choices[0].message.content.strip()

        # Parse the response to extract description and code
        return self._parse_example_response(content)

    def generate_method_example(
        self,
        class_name: str,
        class_description: str,
        method_info: Dict
    ) -> Dict:
        """Generate an example for a method."""
        prompt = self._build_method_example_prompt(
            class_name, class_description, method_info
        )

        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": "You are an expert at creating concise, practical API usage examples."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
        )

        content = response.choices[0].message.content.strip()

        return self._parse_example_response(content)

    def generate_description(
        self,
        class_name: str,
        member_type: str,  # "property" or "method"
        member_name: str,
        usage_code: str
    ) -> str:
        """Generate a task description for existing usage code."""
        prompt = f"""Given this {member_type} usage example for {class_name}.{member_name}, write a brief task requirement that this code accomplishes.

Usage code:
```typescript
{usage_code}
```

Write the task as a concrete requirement (what needs to be done), NOT an explanation of what the code does.

Examples:
- Good: "Set the font size of the paragraph to 14 points"
- Bad: "This code sets the font size of the paragraph to 14 points"

- Good: "Insert a new paragraph at the end of the document"
- Bad: "This example inserts a new paragraph at the end of the document"

Return ONLY the task requirement (1 sentence), no code, no extra formatting."""

        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": "You are an expert at writing concise task requirements."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
        )

        return response.choices[0].message.content.strip()

    def _build_property_example_prompt(
        self,
        class_name: str,
        class_description: str,
        property_info: Dict
    ) -> str:
        """Build prompt for generating a property example."""
        prop_type = property_info.get("type", "unknown")
        prop_desc = property_info.get("description", "")

        prompt = f"""Create a concise TypeScript usage example for the following Word.js API property.

Class: {class_name}
Class Description: {class_description}

Property: {property_info["name"]}
Type: {prop_type}
Description: {prop_desc}

Generate a practical example in task-solution format. Return in this exact format:

DESCRIPTION: [A concrete task/requirement, like "Set the font size to 16" or "Change the border color to red"]

CODE:
```typescript
[TypeScript code implementing the task above using Word.run async pattern]
```

The DESCRIPTION should be a specific task requirement (what to do), and the CODE should implement that task.
Keep it simple and focused on this specific property."""

        return prompt

    def _build_method_example_prompt(
        self,
        class_name: str,
        class_description: str,
        method_info: Dict
    ) -> str:
        """Build prompt for generating a method example."""
        method_desc = method_info.get("description", "")
        signatures = method_info.get("signatures", [])

        # Build signature info
        sig_info = ""
        if signatures:
            sig = signatures[0]  # Use first signature
            params = sig.get("params", [])
            if params:
                param_list = ", ".join([p["name"] for p in params])
                sig_info = f"\nParameters: {param_list}"

        prompt = f"""Create a concise TypeScript usage example for the following Word.js API method.

Class: {class_name}
Class Description: {class_description}

Method: {method_info["name"]}(){sig_info}
Description: {method_desc}

Generate a practical example in task-solution format. Return in this exact format:

DESCRIPTION: [A concrete task/requirement, like "Insert a paragraph with text" or "Delete the first table"]

CODE:
```typescript
[TypeScript code implementing the task above using Word.run async pattern]
```

The DESCRIPTION should be a specific task requirement (what to do), and the CODE should implement that task.
Keep it simple and focused on this specific method."""

        return prompt

    def _parse_example_response(self, content: str) -> Dict:
        """Parse the LLM response into description and code."""
        import re

        # Extract description
        desc_match = re.search(r'DESCRIPTION:\s*(.+?)(?:\n|$)', content, re.IGNORECASE)
        description = desc_match.group(1).strip() if desc_match else None

        # Extract code
        code_match = re.search(r'```(?:typescript|ts)?\n(.*?)```', content, re.DOTALL)
        usage_code = code_match.group(1).strip() if code_match else None

        return {
            "description": description,
            "usage_code": usage_code,
            "output_code": None
        }


class JSONProcessor:
    """Process JSON files to enhance examples."""

    def __init__(
        self,
        generator: ExampleGenerator,
        dry_run: bool = False,
        skip_properties: bool = False,
        skip_methods: bool = False,
        max_per_file: Optional[int] = None
    ):
        self.generator = generator
        self.dry_run = dry_run
        self.skip_properties = skip_properties
        self.skip_methods = skip_methods
        self.max_per_file = max_per_file
        self.stats = {
            "properties_processed": 0,
            "methods_processed": 0,
            "examples_generated": 0,
            "descriptions_generated": 0,
        }

    def process_file(self, json_data: Dict) -> Dict:
        """Process a single JSON file to enhance examples."""
        class_info = json_data.get("class", {})
        class_name = class_info.get("name", "Unknown")
        class_description = class_info.get("description", "")

        generated_count = 0

        # Process properties
        if not self.skip_properties:
            properties = json_data.get("properties", [])
            for prop in properties:
                if self.max_per_file and generated_count >= self.max_per_file:
                    break

                changes = self._process_member_examples(
                    class_name,
                    class_description,
                    prop,
                    "property"
                )

                if changes:
                    generated_count += changes
                    self.stats["properties_processed"] += 1

        # Process methods
        if not self.skip_methods:
            methods = json_data.get("methods", [])
            for method in methods:
                if self.max_per_file and generated_count >= self.max_per_file:
                    break

                changes = self._process_member_examples(
                    class_name,
                    class_description,
                    method,
                    "method"
                )

                if changes:
                    generated_count += changes
                    self.stats["methods_processed"] += 1

        return json_data

    def _process_member_examples(
        self,
        class_name: str,
        class_description: str,
        member: Dict,
        member_type: str
    ) -> int:
        """Process examples for a property or method. Returns number of changes made."""
        member_name = member.get("name", "unknown")
        examples = member.get("examples", [])

        changes = 0

        # Case 1: No examples at all - generate one
        if not examples:
            if self.dry_run:
                print(f"  [DRY] Would generate example for {member_type} '{member_name}'")
            else:
                print(f"  [GENERATE] Creating example for {member_type} '{member_name}'...")
                try:
                    if member_type == "property":
                        new_example = self.generator.generate_property_example(
                            class_name, class_description, member
                        )
                    else:
                        new_example = self.generator.generate_method_example(
                            class_name, class_description, member
                        )

                    member["examples"] = [new_example]
                    self.stats["examples_generated"] += 1
                    changes += 1

                except Exception as exc:
                    print(f"    [ERROR] Failed to generate example: {exc}")

        # Case 2: Has examples - check if any need descriptions
        else:
            for i, example in enumerate(examples):
                usage_code = example.get("usage_code")
                description = example.get("description")

                # Has code but no description
                if usage_code and not description:
                    if self.dry_run:
                        print(f"  [DRY] Would generate description for {member_type} '{member_name}' example {i+1}")
                    else:
                        print(f"  [ENHANCE] Adding description to {member_type} '{member_name}' example {i+1}...")
                        try:
                            new_desc = self.generator.generate_description(
                                class_name, member_type, member_name, usage_code
                            )
                            example["description"] = new_desc
                            self.stats["descriptions_generated"] += 1
                            changes += 1

                        except Exception as exc:
                            print(f"    [ERROR] Failed to generate description: {exc}")

        return changes

    def get_stats(self) -> Dict:
        """Return processing statistics."""
        return self.stats.copy()


def main(argv: Sequence[str] | None = None) -> None:
    args = parse_args(argv)

    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)

    # Validate input directory
    if not input_dir.is_dir():
        raise SystemExit(f"Input directory not found: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)

    # Determine files to process
    files_to_process: List[Path] = []

    if args.file:
        file_path = input_dir / args.file if not Path(args.file).is_absolute() else Path(args.file)
        if not file_path.exists():
            raise SystemExit(f"File not found: {file_path}")
        files_to_process = [file_path]

    elif args.files:
        for filename in args.files:
            file_path = input_dir / filename if not Path(filename).is_absolute() else Path(filename)
            if not file_path.exists():
                print(f"[WARN] File not found: {file_path}", file=sys.stderr)
                continue
            files_to_process.append(file_path)

    elif args.all:
        files_to_process = sorted(input_dir.glob("*.json"))
        if not files_to_process:
            print(f"No JSON files found in {input_dir}")
            return

        print(f"Found {len(files_to_process)} JSON files to process:", file=sys.stderr)
        for f in files_to_process:
            print(f"  - {f.name}", file=sys.stderr)

    if not files_to_process:
        raise SystemExit("No files to process")

    # Initialize OpenAI client
    client = OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key="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",
    )

    # Initialize generator and processor
    generator = ExampleGenerator(client, args.model)
    processor = JSONProcessor(
        generator,
        dry_run=args.dry_run,
        skip_properties=args.skip_properties,
        skip_methods=args.skip_methods,
        max_per_file=args.max_per_file
    )

    # Process files
    successes = 0
    failures = []

    for i, input_path in enumerate(files_to_process, start=1):
        print(f"\n{'='*60}", file=sys.stderr)
        print(f"[{i}/{len(files_to_process)}] Processing {input_path.name}...", file=sys.stderr)

        # Determine output path
        output_path = output_dir / input_path.name

        if output_path.exists() and not args.overwrite:
            print(f"[SKIP] {output_path.name} (already exists)", file=sys.stderr)
            continue

        # Read JSON
        try:
            json_data = json.loads(input_path.read_text(encoding="utf-8"))
        except Exception as exc:
            print(f"[ERROR] Failed to read {input_path.name}: {exc}", file=sys.stderr)
            failures.append(input_path.name)
            continue

        # Process
        try:
            processed_data = processor.process_file(json_data)
        except Exception as exc:
            print(f"[ERROR] Failed to process {input_path.name}: {exc}", file=sys.stderr)
            import traceback
            traceback.print_exc()
            failures.append(input_path.name)
            continue

        # Write output
        if not args.dry_run:
            try:
                output_text = json.dumps(processed_data, ensure_ascii=False, indent=2)
                output_path.write_text(output_text, encoding="utf-8")
                print(f"[OK] Saved to {output_path.name}", file=sys.stderr)
                successes += 1
            except Exception as exc:
                print(f"[ERROR] Failed to write {output_path.name}: {exc}", file=sys.stderr)
                failures.append(input_path.name)
        else:
            successes += 1

    # Summary
    print(f"\n{'='*60}", file=sys.stderr)
    if args.dry_run:
        print(f"Dry-run complete. Would process {len(files_to_process)} files.", file=sys.stderr)
    else:
        print(f"Completed: {successes} successes, {len(failures)} failures", file=sys.stderr)

        stats = processor.get_stats()
        print(f"\nStatistics:", file=sys.stderr)
        print(f"  Properties processed: {stats['properties_processed']}", file=sys.stderr)
        print(f"  Methods processed: {stats['methods_processed']}", file=sys.stderr)
        print(f"  Examples generated: {stats['examples_generated']}", file=sys.stderr)
        print(f"  Descriptions generated: {stats['descriptions_generated']}", file=sys.stderr)

        if failures:
            print(f"\nFailed files:", file=sys.stderr)
            for name in failures:
                print(f"  - {name}", file=sys.stderr)
            raise SystemExit(1)

        print(f"\nAll outputs saved to {output_dir}", file=sys.stderr)


if __name__ == "__main__":
    main()
