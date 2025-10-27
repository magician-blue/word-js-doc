import json
import os
import sys

def format_example(example, indent=0):
    """Format an example with code blocks"""
    lines = []
    indent_str = "  " * indent

    if example.get('description'):
        lines.append(f"{indent_str}**Example**: {example['description']}")
        lines.append("")

    if example.get('usage_code'):
        lines.append(f"{indent_str}```typescript")
        lines.append(example['usage_code'])
        lines.append(f"{indent_str}```")
        lines.append("")

    if example.get('output_code'):
        lines.append(f"{indent_str}**Output**:")
        lines.append(f"{indent_str}```")
        lines.append(example['output_code'])
        lines.append(f"{indent_str}```")
        lines.append("")

    return lines

def format_signature(signature, indent=0):
    """Format a method signature"""
    lines = []
    indent_str = "  " * indent

    # Parameters
    if signature.get('params'):
        lines.append(f"{indent_str}**Parameters:**")
        for param in signature['params']:
            param_name = param.get('name', '')
            param_type = param.get('type', 'any')
            param_required = ' (required)' if param.get('required') else ' (optional)'
            param_desc = param.get('description', '')

            lines.append(f"{indent_str}- `{param_name}`: `{param_type}`{param_required}")
            if param_desc:
                lines.append(f"{indent_str}  {param_desc}")
        lines.append("")

    # Returns
    returns = signature.get('returns')
    if returns:
        return_type = returns.get('type', 'void')
        return_desc = returns.get('description', '')
        lines.append(f"{indent_str}**Returns:** `{return_type}`")
        if return_desc:
            lines.append(f"{indent_str}{return_desc}")
        lines.append("")

    return lines

def json_to_markdown(json_file, output_file=None):
    """Convert a JSON file to Markdown format"""

    # Read JSON file
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    lines = []

    # Class header
    class_info = data.get('class', {})
    class_name = class_info.get('name', 'Unknown Class')

    lines.append(f"# {class_name}")
    lines.append("")

    # Package and API Set
    package = class_info.get('package', '')
    if package:
        lines.append(f"**Package:** `{package}`")
        lines.append("")

    api_set = class_info.get('api_set', {})
    if api_set:
        api_name = api_set.get('name', '')
        api_status = api_set.get('status', '')
        lines.append(f"**API Set:** {api_name} {api_status}")
        lines.append("")

    # Extends
    extends = class_info.get('extends', [])
    if extends:
        lines.append(f"**Extends:** {', '.join([f'`{e}`' for e in extends])}")
        lines.append("")

    # Description
    description = class_info.get('description', '')
    if description:
        lines.append("## Description")
        lines.append("")
        lines.append(description)
        lines.append("")

    # Class-level examples
    class_examples = class_info.get('examples', [])
    if class_examples:
        lines.append("## Class Examples")
        lines.append("")
        for example in class_examples:
            lines.extend(format_example(example))

    # Properties
    properties = data.get('properties', [])
    if properties:
        lines.append("## Properties")
        lines.append("")

        for prop in properties:
            prop_name = prop.get('name', '')
            prop_type = prop.get('type', 'any')
            prop_desc = prop.get('description', '')
            prop_since = prop.get('since', '')

            lines.append(f"### {prop_name}")
            lines.append("")
            lines.append(f"**Type:** `{prop_type}`")
            lines.append("")

            if prop_since:
                lines.append(f"**Since:** {prop_since}")
                lines.append("")

            if prop_desc:
                lines.append(prop_desc)
                lines.append("")

            # Property examples
            prop_examples = prop.get('examples', [])
            if prop_examples:
                lines.append("#### Examples")
                lines.append("")
                for example in prop_examples:
                    lines.extend(format_example(example))

            lines.append("---")
            lines.append("")

    # Methods
    methods = data.get('methods', [])
    if methods:
        lines.append("## Methods")
        lines.append("")

        for method in methods:
            method_name = method.get('name', '')
            method_kind = method.get('kind', '')
            method_desc = method.get('description', '')

            lines.append(f"### {method_name}")
            lines.append("")

            if method_kind:
                lines.append(f"**Kind:** `{method_kind}`")
                lines.append("")

            if method_desc:
                lines.append(method_desc)
                lines.append("")

            # Signatures
            signatures = method.get('signatures', [])
            if signatures:
                if len(signatures) == 1:
                    lines.append("#### Signature")
                    lines.append("")
                    lines.extend(format_signature(signatures[0]))
                else:
                    lines.append("#### Signatures")
                    lines.append("")
                    for i, sig in enumerate(signatures, 1):
                        lines.append(f"**Overload {i}:**")
                        lines.append("")
                        lines.extend(format_signature(sig, indent=1))

            # Method examples
            method_examples = method.get('examples', [])
            if method_examples:
                lines.append("#### Examples")
                lines.append("")
                for example in method_examples:
                    lines.extend(format_example(example))

            lines.append("---")
            lines.append("")

    # Source
    source = data.get('source', {})
    if source:
        urls = source.get('urls', [])
        if urls:
            lines.append("## Source")
            lines.append("")
            for url in urls:
                lines.append(f"- {url}")
            lines.append("")

    # Write to file
    markdown_content = '\n'.join(lines)

    if output_file is None:
        # Generate output filename
        base_name = os.path.splitext(os.path.basename(json_file))[0]
        output_file = base_name + '.md'

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(markdown_content)

    print(f"Successfully converted {json_file} to {output_file}")
    return output_file

def main():
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python itemize.py <json_file> [output_file]")
        print("")
        print("Examples:")
        print("  python itemize.py processed/Word.Border.json")
        print("  python itemize.py processed/Word.Border.json Word.Border.md")
        print("")
        print("Batch conversion:")
        print("  python itemize.py processed/*.json")
        sys.exit(1)

    json_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(json_file):
        print(f"Error: File {json_file} does not exist")
        sys.exit(1)

    json_to_markdown(json_file, output_file)

if __name__ == '__main__':
    main()
