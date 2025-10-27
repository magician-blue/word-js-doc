#!/usr/bin/env python3
"""
Markdown parser for Word.js API documentation.

This module provides utilities to parse large Markdown documentation files
by analyzing their heading hierarchy and extracting structured information.
"""

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple


@dataclass
class MarkdownSection:
    """Represents a section in a Markdown document."""
    level: int  # Heading level (1-6)
    title: str  # Heading text
    content: str  # Content under this heading (excluding sub-sections)
    children: List['MarkdownSection'] = field(default_factory=list)
    line_start: int = 0
    line_end: int = 0


class MarkdownParser:
    """Parse Markdown documents based on heading hierarchy."""

    def __init__(self, text: str):
        self.text = text
        self.lines = text.split('\n')

    def parse(self) -> MarkdownSection:
        """Parse the entire document into a hierarchical structure."""
        root = MarkdownSection(level=0, title="ROOT", content="")

        sections = self._extract_sections()
        root.children = self._build_hierarchy(sections)

        return root

    def _extract_sections(self) -> List[MarkdownSection]:
        """Extract all sections from the document."""
        sections = []
        current_section = None

        for i, line in enumerate(self.lines):
            heading_match = re.match(r'^(#{1,6})\s+(.+)$', line)

            if heading_match:
                # Save previous section
                if current_section:
                    current_section.line_end = i - 1
                    sections.append(current_section)

                # Start new section
                level = len(heading_match.group(1))
                title = heading_match.group(2).strip()
                current_section = MarkdownSection(
                    level=level,
                    title=title,
                    content="",
                    line_start=i
                )
            elif current_section:
                # Add content to current section
                current_section.content += line + '\n'

        # Save last section
        if current_section:
            current_section.line_end = len(self.lines) - 1
            sections.append(current_section)

        return sections

    def _build_hierarchy(self, sections: List[MarkdownSection]) -> List[MarkdownSection]:
        """Build hierarchical structure from flat list of sections."""
        if not sections:
            return []

        root_sections = []
        stack = []

        for section in sections:
            # Pop stack until we find a parent with lower level
            while stack and stack[-1].level >= section.level:
                stack.pop()

            if stack:
                # Add as child to parent
                stack[-1].children.append(section)
            else:
                # Add as root section
                root_sections.append(section)

            stack.append(section)

        return root_sections

    def find_section(self, root: MarkdownSection, title: str) -> Optional[MarkdownSection]:
        """Find a section by title (case-insensitive)."""
        title_lower = title.lower()

        if root.title.lower() == title_lower:
            return root

        for child in root.children:
            result = self.find_section(child, title)
            if result:
                return result

        return None

    def find_sections_by_level(self, root: MarkdownSection, level: int) -> List[MarkdownSection]:
        """Find all sections at a specific heading level."""
        results = []

        if root.level == level:
            results.append(root)

        for child in root.children:
            results.extend(self.find_sections_by_level(child, level))

        return results


class APIDocParser:
    """Parse Word.js API documentation into structured data."""

    def __init__(self, markdown_text: str):
        self.parser = MarkdownParser(markdown_text)
        self.root = self.parser.parse()

    def extract_class_info(self) -> Dict:
        """Extract class-level information."""
        # Find the class title (first h1)
        class_sections = self.parser.find_sections_by_level(self.root, 1)
        if not class_sections:
            return {}

        class_section = class_sections[0]
        class_name = class_section.title.replace(' class', '').strip()

        # Extract package and extends from content
        package = None
        extends = []
        description = None
        api_set = {"name": None, "status": None}

        lines = class_section.content.strip().split('\n')
        for i, line in enumerate(lines):
            if line.startswith('Package:'):
                package_match = re.search(r'Package:\s*\[([^\]]+)\]', line)
                if package_match:
                    package = package_match.group(1)

            if line.startswith('Extends:'):
                extends_match = re.search(r'Extends:\s*\[([^\]]+)\]', line)
                if extends_match:
                    extends.append(extends_match.group(1))

            if i < len(lines) - 1 and not line.startswith('Package:') and not line.startswith('Extends:') and line.strip() and not line.startswith('#'):
                if description is None:
                    description = line.strip()

        # Find Remarks section for API set and examples
        remarks_section = self.parser.find_section(self.root, "Remarks")
        if remarks_section:
            api_match = re.search(r'\[API set:\s*([^\s\]]+)\s*([^\]]*)\]', remarks_section.content)
            if api_match:
                api_set["name"] = api_match.group(1)
                api_set["status"] = api_match.group(2).strip() if api_match.group(2) else None

            # Extract class-level examples from Remarks section
            examples = self._extract_examples(remarks_section.content)

            # Also check for Examples subsection in Remarks
            for subsection in remarks_section.children:
                if subsection.title.lower() == "examples":
                    examples.extend(self._extract_examples(subsection.content))
        else:
            examples = []

        return {
            "name": class_name,
            "package": package,
            "extends": extends,
            "api_set": api_set,
            "description": description,
            "examples": examples
        }

    def extract_properties(self) -> List[Dict]:
        """Extract all properties from Property Details section."""
        properties = []

        prop_details = self.parser.find_section(self.root, "Property Details")
        if not prop_details:
            return properties

        # Each h3 under Property Details is a property
        for prop_section in prop_details.children:
            if prop_section.level != 3:
                continue

            prop_name = prop_section.title.strip()
            prop_data = {
                "name": prop_name,
                "type": None,
                "description": None,
                "since": None,
                "examples": []
            }

            # Extract type from TypeScript code block
            ts_code_match = re.search(r'```(?:typescript|ts)\n(.+?)\n```', prop_section.content, re.DOTALL)
            if ts_code_match:
                ts_code = ts_code_match.group(1).strip()
                # Extract type from TypeScript declaration
                # Matches: "readonly propertyName: Type;" or "propertyName: Type;"
                type_match = re.search(r'(?:readonly\s+)?(?:' + re.escape(prop_name) + r')\s*:\s*([^;]+);?', ts_code)
                if type_match:
                    prop_data["type"] = type_match.group(1).strip()

            # If no TypeScript type found, look for Type: in content
            if not prop_data["type"]:
                content_lines = prop_section.content.strip().split('\n')
                for line in content_lines:
                    if line.startswith('Type:') or re.match(r'^\*\*Type\*\*:', line):
                        type_match = re.search(r'Type\*?\*?:\s*`?([^`\n]+)`?', line)
                        if type_match:
                            prop_data["type"] = type_match.group(1).strip()
                            break

            # Extract description (first non-empty paragraph before code blocks)
            content_lines = prop_section.content.strip().split('\n')
            for line in content_lines:
                # Skip TypeScript code blocks, markdown syntax, and empty lines
                if line.startswith('```') or line.startswith('#') or not line.strip():
                    continue
                if line.startswith('- Property value:') or line.startswith('Remarks'):
                    continue

                if prop_data["description"] is None:
                    # Remove markdown links
                    desc = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', line)
                    prop_data["description"] = desc.strip()
                    break

            # Extract API set from content
            api_match = re.search(r'\[API set:\s*([^\s\]]+)\s*([^\]]*)\]', prop_section.content)
            if api_match:
                prop_data["since"] = f"{api_match.group(1)} {api_match.group(2)}".strip()

            # Extract examples from subsections
            for subsection in prop_section.children:
                if subsection.title.lower() == "examples":
                    prop_data["examples"] = self._extract_examples(subsection.content)

            properties.append(prop_data)

        return properties

    def extract_methods(self) -> List[Dict]:
        """Extract all methods from Method Details section."""
        methods = []
        method_map = {}  # Group overloads by method name

        method_details = self.parser.find_section(self.root, "Method Details")
        if not method_details:
            return methods

        # Each h3 under Method Details is a method
        for method_section in method_details.children:
            if method_section.level != 3:
                continue

            # Parse method signature from title
            method_title = method_section.title.strip()
            method_name, signature = self._parse_method_signature(method_title)

            if method_name not in method_map:
                method_map[method_name] = {
                    "name": method_name,
                    "kind": self._infer_method_kind(method_name),
                    "description": None,
                    "signatures": [],
                    "examples": []
                }

            # Extract description
            content_lines = method_section.content.strip().split('\n')
            for line in content_lines:
                if line.strip() and not line.startswith('#') and not line.startswith('```'):
                    if method_map[method_name]["description"] is None:
                        # Remove markdown links
                        desc = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', line)
                        method_map[method_name]["description"] = desc.strip()
                    break

            # Add signature
            method_map[method_name]["signatures"].append(signature)

            # Extract examples from subsections
            for subsection in method_section.children:
                if subsection.title.lower() == "examples":
                    examples = self._extract_examples(subsection.content)
                    # Only add examples once per method (not per overload)
                    if not method_map[method_name]["examples"]:
                        method_map[method_name]["examples"] = examples

        return list(method_map.values())

    def _parse_method_signature(self, method_title: str) -> Tuple[str, Dict]:
        """Parse method name and signature from title."""
        # Extract method name and parameters
        match = re.match(r'([a-zA-Z0-9_]+)\((.*?)\)', method_title)

        if not match:
            return method_title, {"params": [], "returns": {"type": None, "description": None}}

        method_name = match.group(1)
        params_str = match.group(2).strip()

        # Parse parameters
        params = []
        if params_str:
            # Simple parameter parsing (can be enhanced)
            param_parts = params_str.split(',')
            for param_part in param_parts:
                param_part = param_part.strip()
                if param_part:
                    # Extract parameter name and type
                    param_match = re.match(r'([a-zA-Z0-9_]+)(?:\s*:\s*(.+))?', param_part)
                    if param_match:
                        param_name = param_match.group(1)
                        param_type = param_match.group(2) if param_match.group(2) else None

                        # Check if optional
                        is_optional = '?' in param_part or 'optional' in param_part.lower()

                        params.append({
                            "name": param_name,
                            "type": param_type,
                            "required": not is_optional,
                            "description": None
                        })

        signature = {
            "params": params,
            "returns": {"type": None, "description": None}
        }

        return method_name, signature

    def _infer_method_kind(self, method_name: str) -> Optional[str]:
        """Infer method kind from its name."""
        name_lower = method_name.lower()

        if name_lower.startswith('get'):
            return "read"
        elif name_lower.startswith('set'):
            return "write"
        elif name_lower.startswith('insert') or name_lower.startswith('add'):
            return "create"
        elif name_lower.startswith('delete') or name_lower.startswith('remove') or name_lower == 'clear':
            return "delete"
        elif name_lower == 'load':
            return "load"
        elif name_lower == 'tojson':
            return "serialize"
        elif name_lower == 'track':
            return "track"
        elif name_lower == 'untrack':
            return "untrack"
        else:
            return None

    def _extract_examples(self, content: str) -> List[Dict]:
        """Extract code examples from content."""
        examples = []

        # Find code blocks
        code_blocks = re.findall(r'```(?:[a-zA-Z]*)\n(.*?)```', content, re.DOTALL)

        for code_block in code_blocks:
            code = code_block.strip()

            # Try to find description before the code block
            description = None

            example = {
                "description": description,
                "usage_code": code if code else None,
                "output_code": None
            }
            examples.append(example)

        return examples

    def to_json_schema(self) -> Dict:
        """Convert parsed documentation to JSON schema."""
        return {
            "class": self.extract_class_info(),
            "properties": self.extract_properties(),
            "methods": self.extract_methods(),
            "source": {
                "urls": ["https://docs.microsoft.com/en-us/javascript/api/word"]
            }
        }
