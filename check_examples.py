#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Check all example code snippets in processed/*.json files using parser_optimized.py
Optimized version with persistent TypeScript environment for fast checking.
"""
import os
import sys
import json
import glob
import time
from parser_optimized import check_officejs_ts_fast, cleanup_persistent_env
from typing import Dict, List, Any

# Fix Windows console encoding
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

def collect_examples_from_json(json_file: str) -> List[Dict[str, Any]]:
    """Extract all example code snippets from a JSON file."""
    examples = []

    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    class_name = data.get('class', {}).get('name', 'Unknown')

    # Collect examples from properties
    for prop in data.get('properties', []):
        prop_name = prop.get('name', 'unknown')
        for i, example in enumerate(prop.get('examples', [])):
            usage_code = example.get('usage_code')
            if usage_code:
                examples.append({
                    'type': 'property',
                    'class': class_name,
                    'name': prop_name,
                    'description': example.get('description', ''),
                    'code': usage_code,
                    'location': f"{class_name}.{prop_name} (property, example {i+1})"
                })

    # Collect examples from methods
    for method in data.get('methods', []):
        method_name = method.get('name', 'unknown')
        for i, example in enumerate(method.get('examples', [])):
            usage_code = example.get('usage_code')
            if usage_code:
                examples.append({
                    'type': 'method',
                    'class': class_name,
                    'name': method_name,
                    'description': example.get('description', ''),
                    'code': usage_code,
                    'location': f"{class_name}.{method_name}() (method, example {i+1})"
                })

    return examples


def check_all_processed_files(processed_dir: str = "processed") -> Dict[str, Any]:
    """Check all JSON files in the processed directory."""
    json_files = glob.glob(os.path.join(processed_dir, "*.json"))

    all_results = {
        'total_files': len(json_files),
        'total_examples': 0,
        'passed': 0,
        'failed': 0,
        'files': [],
        'start_time': time.time()
    }

    print(f"Found {len(json_files)} JSON files to check")
    print("Setting up TypeScript environment (this happens only once)...")
    print()

    for file_idx, json_file in enumerate(sorted(json_files), 1):
        filename = os.path.basename(json_file)
        print(f"\n{'='*80}")
        print(f"[{file_idx}/{len(json_files)}] Checking: {filename}")
        print('='*80)

        examples = collect_examples_from_json(json_file)

        file_result = {
            'filename': filename,
            'total_examples': len(examples),
            'passed': 0,
            'failed': 0,
            'failures': []
        }

        for example in examples:
            all_results['total_examples'] += 1

            # Wrap code with type reference and async function wrapper
            code = example['code']

            # Add type reference if not present
            if '/// <reference' not in code:
                code = '/// <reference types="office-js-preview" />\n' + code

            # Wrap in async function if not already wrapped
            if 'async function' not in code.split('\n')[0]:  # Check first line after reference
                # Indent the code
                indented_code = '\n'.join('    ' + line if line.strip() else line
                                         for line in code.split('\n'))
                code = code.split('\n')[0] + '\n' + 'async function func() {\n' + '\n'.join(indented_code.split('\n')[1:]) + '\n}'

            print(f"\n  [{example['type'].upper()}] {example['location']}")
            print(f"  Description: {example['description'][:80]}...")

            # Check the code (using fast version)
            result = check_officejs_ts_fast(code, use_preview=True)

            if result['success']:
                print(f"  ✅ PASS")
                file_result['passed'] += 1
                all_results['passed'] += 1
            else:
                print(f"  ❌ FAIL - {result['error_count']} error(s)")
                print(f"  Errors:")
                for error in result['errors'][:3]:  # Show first 3 errors
                    print(f"    Line {error['line']}:{error['column']} - {error['code']}")
                    print(f"    {error['message'][:100]}")

                file_result['failed'] += 1
                all_results['failed'] += 1

                file_result['failures'].append({
                    'location': example['location'],
                    'description': example['description'],
                    'errors': result['errors']
                })

        all_results['files'].append(file_result)

        # Print file summary
        print(f"\n  File Summary: {file_result['passed']} passed, {file_result['failed']} failed out of {file_result['total_examples']} examples")

    return all_results


def print_summary(results: Dict[str, Any]):
    """Print a final summary of all checks."""
    print("\n" + "="*80)
    print("FINAL SUMMARY")
    print("="*80)

    elapsed = time.time() - results['start_time']

    print(f"\nTotal files checked: {results['total_files']}")
    print(f"Total examples checked: {results['total_examples']}")
    print(f"  ✅ Passed: {results['passed']}")
    print(f"  ❌ Failed: {results['failed']}")

    if results['total_examples'] > 0:
        pass_rate = (results['passed'] / results['total_examples']) * 100
        avg_time = elapsed / results['total_examples']
        print(f"\nPass rate: {pass_rate:.1f}%")
        print(f"Total time: {elapsed:.2f}s")
        print(f"Average time per example: {avg_time:.3f}s")
        print(f"Speed: ~{int(results['total_examples']/elapsed)} examples/second")

    # Show files with failures
    if results['failed'] > 0:
        print("\n" + "-"*80)
        print("Files with failures:")
        print("-"*80)
        for file_result in results['files']:
            if file_result['failed'] > 0:
                print(f"\n  {file_result['filename']}: {file_result['failed']} failure(s)")
                for failure in file_result['failures'][:5]:  # Show first 5 failures per file
                    print(f"    - {failure['location']}")
                    print(f"      {failure['description'][:80]}...")


def save_results_to_json(results: Dict[str, Any], output_file: str = "check_results.json"):
    """Save detailed results to a JSON file."""
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\n\nDetailed results saved to: {output_file}")


if __name__ == "__main__":
    import sys
    import atexit

    # Register cleanup on exit
    atexit.register(cleanup_persistent_env)

    # Check if a specific file is provided
    if len(sys.argv) > 1:
        json_file = sys.argv[1]
        if os.path.exists(json_file):
            start_time = time.time()
            print(f"Checking single file: {json_file}")
            print()
            examples = collect_examples_from_json(json_file)

            passed = 0
            failed = 0

            for idx, example in enumerate(examples, 1):
                code = example['code']

                # Add type reference if not present
                if '/// <reference' not in code:
                    code = '/// <reference types="office-js-preview" />\n' + code

                # Wrap in async function if not already wrapped
                if 'async function' not in code.split('\n')[0]:  # Check first line after reference
                    # Indent the code
                    indented_code = '\n'.join('    ' + line if line.strip() else line
                                             for line in code.split('\n'))
                    code = code.split('\n')[0] + '\n' + 'async function func() {\n' + '\n'.join(indented_code.split('\n')[1:]) + '\n}'

                print(f"\n{'='*80}")
                print(f"[{idx}/{len(examples)}] [{example['type'].upper()}] {example['location']}")
                print(f"Description: {example['description']}")
                print('='*80)

                result = check_officejs_ts_fast(code, use_preview=True)

                if result['success']:
                    print("✅ PASS")
                    passed += 1
                else:
                    print(f"❌ FAIL - {result['error_count']} error(s)")
                    print(f"\n{result['error_summary']}")
                    failed += 1

            elapsed = time.time() - start_time
            print(f"\n{'='*80}")
            print(f"Summary: {passed} passed, {failed} failed out of {len(examples)} examples")
            if len(examples) > 0:
                pass_rate = (passed / len(examples)) * 100
                print(f"Pass rate: {pass_rate:.1f}%")
                print(f"Total time: {elapsed:.2f}s (avg {elapsed/len(examples):.3f}s per example)")
            print('='*80)
        else:
            print(f"Error: File not found: {json_file}")
            sys.exit(1)
    else:
        # Check all files in processed directory
        results = check_all_processed_files()
        print_summary(results)
        save_results_to_json(results)
