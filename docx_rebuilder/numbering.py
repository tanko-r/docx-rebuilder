"""
Numbering Handler - Manages numbering schemes and ensures consistency.

This module handles the detection, normalization, and rebuilding of
numbering schemes to ensure consistent formatting throughout the document.
"""

from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional
from copy import deepcopy

from lxml import etree

from .parser import (
    DocumentStructure, Paragraph, NumberingDefinition, NumberingLevel, NAMESPACES
)


@dataclass
class NumberingState:
    """Tracks the current state of numbering for a scheme."""
    num_id: str
    counters: dict[int, int] = field(default_factory=dict)  # level -> current count

    def increment(self, level: int, start: int = 1) -> int:
        """Increment counter for a level and reset lower levels."""
        # Initialize if needed
        if level not in self.counters:
            self.counters[level] = start - 1

        # Increment this level
        self.counters[level] += 1

        # Reset all lower levels (higher level numbers)
        for l in list(self.counters.keys()):
            if l > level:
                del self.counters[l]

        return self.counters[level]

    def get_current(self, level: int) -> int:
        """Get current counter value for a level."""
        return self.counters.get(level, 0)

    def reset(self, level: int, value: int = 0):
        """Reset a level to a specific value."""
        self.counters[level] = value


@dataclass
class NormalizedNumbering:
    """Normalized numbering definition with clean, consistent values."""
    num_id: str
    abstract_num_id: str
    levels: dict[int, NumberingLevel] = field(default_factory=dict)
    level_indent_base: int = 720  # Base indent per level (twips)


class NumberingHandler:
    """Handles numbering scheme detection and normalization."""

    def __init__(self, structure: DocumentStructure, indent_unit: int = 720):
        self.structure = structure
        self.indent_unit = indent_unit
        self.states: dict[str, NumberingState] = {}
        self.normalized_definitions: dict[str, NormalizedNumbering] = {}

    def analyze_and_normalize(self) -> dict[str, NormalizedNumbering]:
        """Analyze numbering schemes and create normalized versions."""
        for num_id, num_def in self.structure.numbering_definitions.items():
            normalized = self._normalize_definition(num_def)
            self.normalized_definitions[num_id] = normalized

        return self.normalized_definitions

    def _normalize_definition(self, num_def: NumberingDefinition) -> NormalizedNumbering:
        """Create a normalized version of a numbering definition."""
        normalized = NormalizedNumbering(
            num_id=num_def.num_id,
            abstract_num_id=num_def.abstract_num_id,
            level_indent_base=self.indent_unit
        )

        for level_num, level_def in num_def.levels.items():
            # Normalize indent to be consistent with level
            expected_indent = (level_num + 1) * self.indent_unit

            # Calculate hanging indent (typically equal to indent unit)
            hanging = self.indent_unit

            normalized_level = NumberingLevel(
                level=level_num,
                start=level_def.start,
                num_fmt=level_def.num_fmt,
                level_text=level_def.level_text,
                indent_left=expected_indent,
                indent_hanging=hanging,
                alignment=level_def.alignment
            )
            normalized.levels[level_num] = normalized_level

        return normalized

    def get_numbering_text(self, num_id: str, level: int) -> str:
        """Get the formatted numbering text for a paragraph."""
        if num_id not in self.states:
            self.states[num_id] = NumberingState(num_id=num_id)

        state = self.states[num_id]
        num_def = self.structure.numbering_definitions.get(num_id)

        if not num_def or level not in num_def.levels:
            return ""

        level_def = num_def.levels[level]

        # Increment counter
        start = level_def.start
        current = state.increment(level, start)

        # Format the number
        formatted = self._format_number(current, level_def.num_fmt)

        # Apply level text pattern
        text = level_def.level_text
        for i in range(level + 1):
            placeholder = f"%{i + 1}"
            if placeholder in text:
                level_num = state.get_current(i) if i < level else current
                level_formatted = self._format_number(level_num, num_def.levels.get(i, level_def).num_fmt)
                text = text.replace(placeholder, level_formatted)

        return text

    def _format_number(self, num: int, fmt: str) -> str:
        """Format a number according to the numbering format."""
        if fmt == 'decimal':
            return str(num)
        elif fmt == 'lowerLetter':
            return self._to_letter(num, lowercase=True)
        elif fmt == 'upperLetter':
            return self._to_letter(num, lowercase=False)
        elif fmt == 'lowerRoman':
            return self._to_roman(num).lower()
        elif fmt == 'upperRoman':
            return self._to_roman(num)
        elif fmt == 'bullet':
            return '•'
        elif fmt == 'none':
            return ''
        else:
            return str(num)

    def _to_letter(self, num: int, lowercase: bool = True) -> str:
        """Convert number to letter (1=a, 2=b, etc.)."""
        if num <= 0:
            return ''

        result = []
        while num > 0:
            num -= 1
            result.append(chr((num % 26) + (ord('a') if lowercase else ord('A'))))
            num //= 26

        return ''.join(reversed(result))

    def _to_roman(self, num: int) -> str:
        """Convert number to Roman numerals."""
        if num <= 0:
            return ''

        val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
        syms = ['M', 'CM', 'D', 'CD', 'C', 'XC', 'L', 'XL', 'X', 'IX', 'V', 'IV', 'I']

        result = ''
        for i, v in enumerate(val):
            while num >= v:
                result += syms[i]
                num -= v

        return result

    def reset_all_states(self):
        """Reset all numbering states."""
        self.states.clear()

    def reset_state(self, num_id: str):
        """Reset the state for a specific numbering scheme."""
        if num_id in self.states:
            del self.states[num_id]

    def build_numbering_xml(self) -> etree._Element:
        """Build clean numbering.xml content."""
        nsmap = {
            'w': NAMESPACES['w'],
            'r': NAMESPACES['r'],
        }

        root = etree.Element(f'{{{NAMESPACES["w"]}}}numbering', nsmap=nsmap)

        # Build abstract numbering definitions
        abstract_nums_created = set()

        for num_id, normalized in self.normalized_definitions.items():
            if normalized.abstract_num_id not in abstract_nums_created:
                abstract_elem = self._build_abstract_num(normalized)
                root.append(abstract_elem)
                abstract_nums_created.add(normalized.abstract_num_id)

        # Build numbering instances
        for num_id, normalized in self.normalized_definitions.items():
            num_elem = etree.SubElement(root, f'{{{NAMESPACES["w"]}}}num')
            num_elem.set(f'{{{NAMESPACES["w"]}}}numId', num_id)

            abstract_ref = etree.SubElement(num_elem, f'{{{NAMESPACES["w"]}}}abstractNumId')
            abstract_ref.set(f'{{{NAMESPACES["w"]}}}val', normalized.abstract_num_id)

        return root

    def _build_abstract_num(self, normalized: NormalizedNumbering) -> etree._Element:
        """Build an abstract numbering definition element."""
        abstract_elem = etree.Element(f'{{{NAMESPACES["w"]}}}abstractNum')
        abstract_elem.set(f'{{{NAMESPACES["w"]}}}abstractNumId', normalized.abstract_num_id)

        # Add multiLevelType
        multi_level = etree.SubElement(abstract_elem, f'{{{NAMESPACES["w"]}}}multiLevelType')
        multi_level.set(f'{{{NAMESPACES["w"]}}}val', 'multilevel')

        # Add levels
        for level_num in sorted(normalized.levels.keys()):
            level_def = normalized.levels[level_num]
            lvl_elem = self._build_level_element(level_def)
            abstract_elem.append(lvl_elem)

        return abstract_elem

    def _build_level_element(self, level_def: NumberingLevel) -> etree._Element:
        """Build a level element for numbering."""
        lvl = etree.Element(f'{{{NAMESPACES["w"]}}}lvl')
        lvl.set(f'{{{NAMESPACES["w"]}}}ilvl', str(level_def.level))

        # Start value
        start = etree.SubElement(lvl, f'{{{NAMESPACES["w"]}}}start')
        start.set(f'{{{NAMESPACES["w"]}}}val', str(level_def.start))

        # Number format
        num_fmt = etree.SubElement(lvl, f'{{{NAMESPACES["w"]}}}numFmt')
        num_fmt.set(f'{{{NAMESPACES["w"]}}}val', level_def.num_fmt)

        # Level text
        lvl_text = etree.SubElement(lvl, f'{{{NAMESPACES["w"]}}}lvlText')
        lvl_text.set(f'{{{NAMESPACES["w"]}}}val', level_def.level_text)

        # Level justification
        lvl_jc = etree.SubElement(lvl, f'{{{NAMESPACES["w"]}}}lvlJc')
        lvl_jc.set(f'{{{NAMESPACES["w"]}}}val', level_def.alignment)

        # Paragraph properties (indentation)
        pPr = etree.SubElement(lvl, f'{{{NAMESPACES["w"]}}}pPr')
        ind = etree.SubElement(pPr, f'{{{NAMESPACES["w"]}}}ind')

        if level_def.indent_left:
            ind.set(f'{{{NAMESPACES["w"]}}}left', str(level_def.indent_left))
        if level_def.indent_hanging:
            ind.set(f'{{{NAMESPACES["w"]}}}hanging', str(level_def.indent_hanging))

        return lvl


def detect_numbering_pattern(paragraphs: list[Paragraph]) -> list[dict]:
    """Detect numbering patterns from paragraph text content.

    This is useful for documents where numbering isn't properly defined
    in numbering.xml but appears as plain text.
    """
    import re

    patterns = [
        # Numeric: 1. 2. 3. or 1) 2) 3)
        (r'^(\d+)[.\)]\s+', 'decimal'),
        # Letters: a. b. c. or a) b) c)
        (r'^([a-z])[.\)]\s+', 'lowerLetter'),
        (r'^([A-Z])[.\)]\s+', 'upperLetter'),
        # Roman: i. ii. iii. or (i) (ii)
        (r'^([ivxlcdm]+)[.\)]\s+', 'lowerRoman'),
        (r'^([IVXLCDM]+)[.\)]\s+', 'upperRoman'),
        # Bullets
        (r'^[•●○▪▸►]\s+', 'bullet'),
        # Section numbers: 1.1, 1.1.1, etc.
        (r'^(\d+(?:\.\d+)+)[.\s]+', 'multilevel'),
    ]

    detected = []

    for i, para in enumerate(paragraphs):
        text = para.text.strip()

        for pattern, fmt in patterns:
            match = re.match(pattern, text)
            if match:
                detected.append({
                    'paragraph_index': i,
                    'format': fmt,
                    'match': match.group(0),
                    'value': match.group(1) if match.groups() else None
                })
                break

    return detected
