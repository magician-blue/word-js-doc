"""
Analyzes usage patterns between Classes, Interfaces, and Enums in Word.js documentation.
"""

import re
from typing import Set, Dict, Tuple, List


def parse_markdown_table(file_path: str) -> List[Tuple[str, str]]:
    """
    Parse markdown table and extract name-description pairs.

    Args:
        file_path: Path to markdown file

    Returns:
        List of tuples containing (name, description)
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Extract table rows (skip header and separator)
    lines = content.strip().split('\n')
    items = []

    for line in lines:
        # Skip headers, separators, and empty lines
        if line.startswith('|---') or line.startswith('## ') or not line.strip() or line.startswith('| Enum |') or line.startswith('| Class |') or line.startswith('| Interface |'):
            continue

        # Parse table row
        if line.startswith('|'):
            parts = [p.strip() for p in line.split('|')]
            #parts[0] is empty, parts[1] is name, parts[2] is description
            if len(parts) >= 3 and parts[1]:
                items.append((parts[1], parts[2] if len(parts) > 2 else ''))

    return items


def analyze_class_interface_overlap(class_md: str, interface_md: str) -> Dict:
    """
    Analyzes overlap between Classes and Word.Interfaces.

    This function checks how many classes have corresponding interfaces
    in the Word.Interfaces namespace.

    Args:
        class_md: Path to class.md file
        interface_md: Path to interface.md file

    Returns:
        Dictionary containing:
        - total_classes: Total number of classes
        - total_interfaces: Total number of Word.Interfaces items
        - classes_with_interfaces: Set of class names that have corresponding interfaces
        - overlap_count: Number of classes with corresponding interfaces
        - overlap_percentage: Percentage of overlap
    """
    # Parse files
    classes = parse_markdown_table(class_md)
    interfaces = parse_markdown_table(interface_md)

    # Extract class names (remove Word. prefix)
    class_names = set()
    for name, desc in classes:
        # Extract just the class name without Word. prefix
        if name.startswith('Word.'):
            class_name = name.replace('Word.', '')
            class_names.add(class_name)

    # Extract Word.Interfaces items
    word_interfaces = set()
    for name, desc in interfaces:
        if name.startswith('Word.Interfaces.'):
            # Extract the interface name
            interface_name = name.replace('Word.Interfaces.', '')
            word_interfaces.add(interface_name)

    # Find classes that have corresponding interfaces
    # Look for patterns like: ClassName -> ClassNameData, ClassNameLoadOptions, ClassNameUpdateData
    classes_with_interfaces = set()

    for class_name in class_names:
        # Check if any Word.Interfaces item starts with this class name
        for interface_name in word_interfaces:
            if (interface_name.startswith(class_name + 'Data') or
                interface_name.startswith(class_name + 'LoadOptions') or
                interface_name.startswith(class_name + 'UpdateData') or
                interface_name.startswith(class_name + 'CollectionData') or
                interface_name.startswith(class_name + 'CollectionLoadOptions') or
                interface_name.startswith(class_name + 'CollectionUpdateData')):
                classes_with_interfaces.add(class_name)
                break

    overlap_count = len(classes_with_interfaces)
    overlap_percentage = (overlap_count / len(class_names) * 100) if class_names else 0

    return {
        'total_classes': len(class_names),
        'total_word_interfaces': len(word_interfaces),
        'total_all_interfaces': len(interfaces),
        'classes_with_interfaces': classes_with_interfaces,
        'overlap_count': overlap_count,
        'overlap_percentage': overlap_percentage,
        'classes_without_interfaces': class_names - classes_with_interfaces
    }


def analyze_enum_usage_in_interfaces(enum_md: str, interface_md: str, api_docs_dir: str = None) -> Dict:
    """
    Analyzes how many enums are used in interface descriptions.

    This function extracts all enums from enums.md and searches for their
    usage in interface descriptions and corresponding API documentation files.

    Args:
        enum_md: Path to enums.md file
        interface_md: Path to interface.md file
        api_docs_dir: Optional path to directory containing API markdown files

    Returns:
        Dictionary containing:
        - total_enums: Total number of enums
        - used_enums: Set of enum names found in interface descriptions
        - used_count: Number of enums used
        - usage_percentage: Percentage of enums used
        - unused_enums: Set of enums not found in interfaces
    """
    import os
    from pathlib import Path

    # Parse files
    enums = parse_markdown_table(enum_md)
    interfaces = parse_markdown_table(interface_md)

    # Extract enum names
    enum_names = set()
    for name, desc in enums:
        enum_names.add(name)

    # Extract interface names
    interface_names = []
    for name, desc in interfaces:
        interface_names.append(name)

    # Search in API documentation markdown files for each interface
    api_docs_text = ""
    files_found = 0
    files_not_found = 0

    if api_docs_dir and os.path.exists(api_docs_dir):
        for interface_name in interface_names:
            # Construct the file path
            api_file = Path(api_docs_dir) / f"{interface_name}.md"

            if api_file.exists():
                try:
                    with open(api_file, 'r', encoding='utf-8') as f:
                        api_docs_text += " " + f.read()
                    files_found += 1
                except Exception as e:
                    # Skip files with errors silently
                    pass
            else:
                files_not_found += 1

        print(f"Found {files_found} API docs for interfaces, {files_not_found} not found")

    # Find which enums are mentioned in interface docs
    used_enums = set()
    enum_usage_details = {}

    for enum_name in enum_names:
        # Search for the enum name in combined text
        # Use word boundary to avoid partial matches
        pattern = r'\b' + re.escape(enum_name) + r'\b'
        matches = re.findall(pattern, api_docs_text)

        if matches:
            used_enums.add(enum_name)
            enum_usage_details[enum_name] = len(matches)

    used_count = len(used_enums)
    usage_percentage = (used_count / len(enum_names) * 100) if enum_names else 0

    return {
        'total_enums': len(enum_names),
        'used_enums': used_enums,
        'used_count': used_count,
        'usage_percentage': usage_percentage,
        'unused_enums': enum_names - used_enums,
        'usage_details': enum_usage_details
    }


def analyze_enum_usage_in_classes(enum_md: str, class_md: str, api_docs_dir: str = None) -> Dict:
    """
    Analyzes how many enums are used in class descriptions and API documentation.

    This function extracts all enums from enums.md and searches for their
    usage in class descriptions and corresponding API documentation files.

    Args:
        enum_md: Path to enums.md file
        class_md: Path to class.md file
        api_docs_dir: Optional path to directory containing API markdown files

    Returns:
        Dictionary containing:
        - total_enums: Total number of enums
        - used_enums: Set of enum names found in class descriptions
        - used_count: Number of enums used
        - usage_percentage: Percentage of enums used
        - unused_enums: Set of enums not found in classes
        - usage_details: Dictionary mapping enum names to usage counts
    """
    import os
    from pathlib import Path

    # Parse files
    enums = parse_markdown_table(enum_md)
    classes = parse_markdown_table(class_md)

    # Extract enum names
    enum_names = set()
    for name, desc in enums:
        enum_names.add(name)

    # Extract class names
    class_names = []
    for name, desc in classes:
        class_names.append(name)

    # Search in API documentation markdown files for each class
    api_docs_text = ""
    files_found = 0
    files_not_found = 0

    if api_docs_dir and os.path.exists(api_docs_dir):
        for class_name in class_names:
            # Construct the file path
            api_file = Path(api_docs_dir) / f"{class_name}.md"

            if api_file.exists():
                try:
                    with open(api_file, 'r', encoding='utf-8') as f:
                        api_docs_text += " " + f.read()
                    files_found += 1
                except Exception as e:
                    # Skip files with errors silently
                    pass
            else:
                files_not_found += 1

        print(f"Found {files_found} API docs for classes, {files_not_found} not found")

    # Find which enums are mentioned in class docs
    used_enums = set()
    enum_usage_details = {}

    for enum_name in enum_names:
        # Search for the enum name in combined text
        # Use word boundary to avoid partial matches
        pattern = r'\b' + re.escape(enum_name) + r'\b'
        matches = re.findall(pattern, api_docs_text)

        if matches:
            used_enums.add(enum_name)
            enum_usage_details[enum_name] = len(matches)

    used_count = len(used_enums)
    usage_percentage = (used_count / len(enum_names) * 100) if enum_names else 0

    return {
        'total_enums': len(enum_names),
        'used_enums': used_enums,
        'used_count': used_count,
        'usage_percentage': usage_percentage,
        'unused_enums': enum_names - used_enums,
        'usage_details': enum_usage_details
    }


if __name__ == '__main__':
    import os

    # Example usage
    class_md_path = 'class.md'
    interface_md_path = 'interface.md'
    enum_md_path = 'enums.md'

    print("=" * 80)
    print("ANALYSIS 1: Class-Interface Overlap")
    print("=" * 80)
    result1 = analyze_class_interface_overlap(class_md_path, interface_md_path)

    print(f"\nTotal Classes: {result1['total_classes']}")
    print(f"Total Word.Interfaces items: {result1['total_word_interfaces']}")
    print(f"Total All Interfaces: {result1['total_all_interfaces']}")
    print(f"Classes with corresponding interfaces: {result1['overlap_count']}")
    print(f"Overlap percentage: {result1['overlap_percentage']:.2f}%")

    print(f"\nClasses WITH interfaces ({len(result1['classes_with_interfaces'])}):")
    for class_name in sorted(result1['classes_with_interfaces'])[:10]:
        print(f"  - {class_name}")
    if len(result1['classes_with_interfaces']) > 10:
        print(f"  ... and {len(result1['classes_with_interfaces']) - 10} more")

    print(f"\nClasses WITHOUT interfaces ({len(result1['classes_without_interfaces'])}):")
    for class_name in sorted(result1['classes_without_interfaces']):
        print(f"  - {class_name}")

    print("\n" + "=" * 80)
    print("ANALYSIS 2: Enum Usage in Interfaces and API Docs")
    print("=" * 80)

    # Try to find the api_docs directory
    api_docs_dir = 'api_docs'
    if not os.path.exists(api_docs_dir):
        api_docs_dir = None
        print("Warning: api_docs directory not found, searching only in interface.md")

    result2 = analyze_enum_usage_in_interfaces(enum_md_path, interface_md_path, api_docs_dir)

    print(f"\nTotal Enums: {result2['total_enums']}")
    print(f"Enums used in interfaces/API docs: {result2['used_count']}")
    print(f"Usage percentage: {result2['usage_percentage']:.2f}%")

    print(f"\nEnums USED in interfaces/API docs ({len(result2['used_enums'])}):")
    # Sort by usage count
    sorted_used = sorted(result2['usage_details'].items(), key=lambda x: x[1], reverse=True)
    for enum_name, count in sorted_used[:20]:
        print(f"  - {enum_name} (used {count} times)")
    if len(sorted_used) > 20:
        print(f"  ... and {len(sorted_used) - 20} more")

    print(f"\nEnums NOT USED in interfaces/API docs ({len(result2['unused_enums'])}):")
    for enum_name in sorted(result2['unused_enums'])[:20]:
        print(f"  - {enum_name}")
    if len(result2['unused_enums']) > 20:
        print(f"  ... and {len(result2['unused_enums']) - 20} more")

    print("\n" + "=" * 80)
    print("ANALYSIS 3: Enum Usage in Classes and API Docs")
    print("=" * 80)

    result3 = analyze_enum_usage_in_classes(enum_md_path, class_md_path, api_docs_dir)

    print(f"\nTotal Enums: {result3['total_enums']}")
    print(f"Enums used in classes/API docs: {result3['used_count']}")
    print(f"Usage percentage: {result3['usage_percentage']:.2f}%")

    print(f"\nEnums USED in classes/API docs ({len(result3['used_enums'])}):")
    # Sort by usage count
    sorted_used = sorted(result3['usage_details'].items(), key=lambda x: x[1], reverse=True)
    for enum_name, count in sorted_used[:20]:
        print(f"  - {enum_name} (used {count} times)")
    if len(sorted_used) > 20:
        print(f"  ... and {len(sorted_used) - 20} more")

    print(f"\nEnums NOT USED in classes/API docs ({len(result3['unused_enums'])}):")
    for enum_name in sorted(result3['unused_enums'])[:20]:
        print(f"  - {enum_name}")
    if len(result3['unused_enums']) > 20:
        print(f"  ... and {len(result3['unused_enums']) - 20} more")

    # Comparison summary
    print("\n" + "=" * 80)
    print("COMPARISON SUMMARY")
    print("=" * 80)
    print(f"\nEnum usage in Interfaces: {result2['usage_percentage']:.2f}% ({result2['used_count']}/{result2['total_enums']})")
    print(f"Enum usage in Classes:    {result3['usage_percentage']:.2f}% ({result3['used_count']}/{result3['total_enums']})")

    # Find enums used in both
    used_in_both = result2['used_enums'] & result3['used_enums']
    only_in_interfaces = result2['used_enums'] - result3['used_enums']
    only_in_classes = result3['used_enums'] - result2['used_enums']

    print(f"\nEnums used in BOTH: {len(used_in_both)}")
    print(f"Enums used ONLY in Interfaces: {len(only_in_interfaces)}")
    print(f"Enums used ONLY in Classes: {len(only_in_classes)}")

    if only_in_interfaces:
        print(f"\nEnums ONLY in Interfaces ({len(only_in_interfaces)}):")
        for enum_name in sorted(only_in_interfaces)[:10]:
            print(f"  - {enum_name}")
        if len(only_in_interfaces) > 10:
            print(f"  ... and {len(only_in_interfaces) - 10} more")

    if only_in_classes:
        print(f"\nEnums ONLY in Classes ({len(only_in_classes)}):")
        for enum_name in sorted(only_in_classes)[:10]:
            print(f"  - {enum_name}")
        if len(only_in_classes) > 10:
            print(f"  ... and {len(only_in_classes) - 10} more")
