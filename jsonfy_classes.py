#!/usr/bin/env python3
"""
Batch converts class documentation Markdown into JSON using the GPT-5 API.

For every class listed in `class.md`, the script reads the matching file inside
`api_docs/`, injects it into the prompt template stored in `jsonfy.md`, and
asks the GPT-5 model to return JSON that complies with the schema described
there. Results are written to `jsonfied/<ClassName>.json`.

Prerequisites:
  * `pip install openai`
  * Set the `OPENAI_API_KEY` environment variable (or configure the OpenAI SDK
    otherwise).
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Sequence

try:
    from openai import OpenAI
except ModuleNotFoundError as exc:  # pragma: no cover - gives a clean hint if deps are missing
    raise SystemExit(
        "Missing dependency: install the OpenAI SDK first, e.g. `pip install openai`."
    ) from exc


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert Word JS class docs in api_docs/ to JSON using GPT-5."
    )
    parser.add_argument(
        "--class-table",
        default="class.md",
        help="Markdown table that lists all classes (default: class.md)",
    )
    parser.add_argument(
        "--docs-dir",
        default="api_docs",
        help="Directory containing Markdown files per class (default: api_docs)",
    )
    parser.add_argument(
        "--prompt-template",
        default="jsonfy.md",
        help="Markdown file that contains the JSON conversion instructions (default: jsonfy.md)",
    )
    parser.add_argument(
        "--output-dir",
        default="jsonfied",
        help="Directory to write generated JSON files into (default: jsonfied)",
    )
    parser.add_argument(
        "--model",
        default="gpt-5.0-mini",
        help="GPT-5 family model name to use (default: gpt-5.0-mini)",
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="Only process the first N classes (useful for smoke tests)",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing JSON files instead of skipping them",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print what would be done without calling the API or writing files",
    )
    return parser.parse_args(argv)


def read_class_names(table_path: Path) -> List[str]:
    classes: List[str] = []
    for raw_line in table_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line.startswith("|"):
            continue
        parts = [part.strip() for part in line.split("|")]
        if len(parts) < 3:
            continue
        candidate = parts[1]
        if not candidate or candidate.lower() == "class":
            continue
        if set(candidate) == {"-"}:
            continue
        classes.append(candidate)
    # Preserve order but drop accidental duplicates.
    seen: Dict[str, None] = {}
    return [seen.setdefault(name, None) or name for name in classes if name not in seen]


def build_messages(prompt_template: str, doc_text: str) -> List[Dict[str, str]]:
    """
    Convert the human-readable template into messages for the Responses API.
    """
    injected = prompt_template.replace("[PASTE CLASS MARKDOWN HERE]", doc_text)
    system_hint = "You are an expert API documentation structurer."
    messages: List[Dict[str, str]] = []
    if system_hint in injected:
        injected = injected.replace(system_hint, "", 1).strip()
        messages.append({"role": "system", "content": system_hint})
    else:
        messages.append({"role": "system", "content": system_hint})
    messages.append({"role": "user", "content": injected})
    return messages


def extract_response_text(response) -> str:
    """
    Normalises the Responses API payload into a raw text string.
    """
    if hasattr(response, "output_text") and response.output_text:
        return response.output_text

    output = getattr(response, "output", None)
    chunks: List[str] = []
    if output:
        for item in output:
            content = getattr(item, "content", None)
            if content is None and isinstance(item, dict):
                content = item.get("content")
            if not content:
                continue
            for block in content:
                block_type = getattr(block, "type", None)
                block_text = getattr(block, "text", None)
                if isinstance(block, dict):
                    block_type = block.get("type")
                    block_text = block.get("text")
                if block_type in {"text", "output_text"} and block_text:
                    chunks.append(block_text)
    if chunks:
        return "".join(chunks)

    choices = getattr(response, "choices", None)
    if choices:
        first = choices[0]
        message = getattr(first, "message", None)
        if message is None and isinstance(first, dict):
            message = first.get("message")
        if message:
            content = getattr(message, "content", None)
            if content is None and isinstance(message, dict):
                content = message.get("content")
            if content:
                if isinstance(content, str):
                    return content
                if isinstance(content, list):
                    return "".join(
                        part.get("text", "")
                        for part in content
                        if isinstance(part, dict) and part.get("type") == "text"
                    )
    raise RuntimeError("Could not extract text from the model response payload.")


def call_model(client: OpenAI, messages: Iterable[Dict[str, str]], model: str) -> str:
    response = client.responses.create(
        model=model,
        input=list(messages),
        temperature=0,
    )
    return extract_response_text(response)


def main(argv: Sequence[str] | None = None) -> None:
    args = parse_args(argv)
    table_path = Path(args.class_table)
    docs_dir = Path(args.docs_dir)
    prompt_path = Path(args.prompt_template)
    output_dir = Path(args.output_dir)

    if not table_path.is_file():
        raise SystemExit(f"Class table not found: {table_path}")
    if not docs_dir.is_dir():
        raise SystemExit(f"Docs directory not found: {docs_dir}")
    if not prompt_path.is_file():
        raise SystemExit(f"Prompt template not found: {prompt_path}")

    output_dir.mkdir(parents=True, exist_ok=True)

    prompt_template = prompt_path.read_text(encoding="utf-8")
    class_names = read_class_names(table_path)
    if args.limit is not None:
        class_names = class_names[: args.limit]

    if not class_names:
        raise SystemExit("No classes found in the supplied table.")

    client = OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key="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",
    )

    successes = 0
    failures: List[str] = []

    for index, class_name in enumerate(class_names, start=1):
        doc_path = docs_dir / f"{class_name}.md"
        if not doc_path.exists():
            print(
                f"[WARN] ({index}/{len(class_names)}) Missing documentation file for {class_name}: {doc_path}",
                file=sys.stderr,
            )
            failures.append(class_name)
            continue

        output_path = output_dir / f"{class_name}.json"
        if output_path.exists() and not args.overwrite:
            print(
                f"[SKIP] ({index}/{len(class_names)}) {class_name} -> {output_path} (already exists)",
                file=sys.stderr,
            )
            continue

        doc_text = doc_path.read_text(encoding="utf-8")
        messages = build_messages(prompt_template, doc_text)

        if args.dry_run:
            print(f"[DRY] ({index}/{len(class_names)}) Would process {class_name}", file=sys.stderr)
            continue

        try:
            raw_text = call_model(client, messages, args.model)
        except Exception as exc:  # noqa: BLE001 - surface the failure with context
            print(
                f"[ERROR] ({index}/{len(class_names)}) OpenAI request failed for {class_name}: {exc}",
                file=sys.stderr,
            )
            failures.append(class_name)
            continue

        try:
            parsed = json.loads(raw_text)
        except json.JSONDecodeError as exc:
            print(
                f"[ERROR] ({index}/{len(class_names)}) Invalid JSON for {class_name}: {exc}",
                file=sys.stderr,
            )
            recovery_path = output_path.with_suffix(".raw.json")
            recovery_path.write_text(raw_text, encoding="utf-8")
            print(
                f"        Wrote raw output for manual inspection: {recovery_path}",
                file=sys.stderr,
            )
            failures.append(class_name)
            continue

        formatted = json.dumps(parsed, ensure_ascii=False, indent=2)
        output_path.write_text(formatted, encoding="utf-8")
        print(
            f"[OK] ({index}/{len(class_names)}) {class_name} -> {output_path}",
            file=sys.stderr,
        )
        successes += 1

    if args.dry_run:
        print(
            f"Dry-run complete. Planned {len(class_names)} classes, skipped API calls.",
            file=sys.stderr,
        )
        return

    if failures:
        print(
            f"Finished with {successes} successes and {len(failures)} failures.",
            file=sys.stderr,
        )
        for name in failures:
            print(f"  - {name}", file=sys.stderr)
        raise SystemExit(1)

    print(
        f"Completed {successes} classes successfully. Outputs saved to {output_dir}.",
        file=sys.stderr,
    )


if __name__ == "__main__":
    main()
