"""
Cross-Reference Handler - Preserves and rebuilds dynamic cross-references.

This module handles the preservation and rebuilding of cross-references,
bookmarks, and field codes to maintain document integrity.
"""

from dataclasses import dataclass, field
from typing import Optional
from lxml import etree
from copy import deepcopy

from .parser import (
    DocumentStructure, CrossReference, Bookmark, Paragraph, NAMESPACES
)


@dataclass
class FieldDefinition:
    """Represents a field (cross-reference, page number, etc.)."""
    field_type: str  # REF, PAGEREF, SEQ, TOC, etc.
    field_code: str
    bookmark_name: Optional[str] = None
    switches: list[str] = field(default_factory=list)
    preserved_formatting: bool = True


@dataclass
class BookmarkMapping:
    """Maps old bookmark names to new ones if renamed during rebuild."""
    original_name: str
    new_name: str
    bookmark_id: str


class CrossReferenceHandler:
    """Handles cross-reference preservation and rebuilding."""

    def __init__(self, structure: DocumentStructure):
        self.structure = structure
        self.bookmark_mappings: dict[str, BookmarkMapping] = {}
        self.field_definitions: list[FieldDefinition] = []
        self.next_bookmark_id = 0

    def analyze_references(self):
        """Analyze all cross-references and bookmarks in the document."""
        # Find the highest existing bookmark ID
        for bookmark in self.structure.bookmarks.values():
            try:
                bm_id = int(bookmark.bookmark_id)
                if bm_id >= self.next_bookmark_id:
                    self.next_bookmark_id = bm_id + 1
            except ValueError:
                pass

        # Parse cross-references into field definitions
        for cross_ref in self.structure.cross_references:
            field_def = self._parse_field_code(cross_ref.field_code)
            if field_def:
                self.field_definitions.append(field_def)

    def _parse_field_code(self, code: str) -> Optional[FieldDefinition]:
        """Parse a field code string into a FieldDefinition."""
        parts = code.strip().split()
        if not parts:
            return None

        field_type = parts[0].upper()
        bookmark_name = None
        switches = []

        i = 1
        while i < len(parts):
            part = parts[i]
            if part.startswith('\\'):
                # This is a switch
                switches.append(part)
            elif bookmark_name is None and not part.startswith('\\'):
                # This is the bookmark name
                bookmark_name = part
            i += 1

        return FieldDefinition(
            field_type=field_type,
            field_code=code,
            bookmark_name=bookmark_name,
            switches=switches,
            preserved_formatting='\\* MERGEFORMAT' in code or '\\h' in switches
        )

    def validate_references(self) -> list[dict]:
        """Validate that all cross-references point to existing bookmarks."""
        issues = []

        for field_def in self.field_definitions:
            if field_def.bookmark_name and field_def.field_type in ('REF', 'PAGEREF'):
                if field_def.bookmark_name not in self.structure.bookmarks:
                    issues.append({
                        'type': 'missing_bookmark',
                        'field_code': field_def.field_code,
                        'bookmark_name': field_def.bookmark_name,
                        'message': f'Cross-reference to non-existent bookmark: {field_def.bookmark_name}'
                    })

        return issues

    def build_field_xml(self, field_def: FieldDefinition, display_text: str = "") -> list[etree._Element]:
        """Build XML elements for a field (cross-reference).

        Returns a list of run elements that make up the complete field.
        """
        runs = []

        # Field begin
        begin_run = etree.Element(f'{{{NAMESPACES["w"]}}}r')
        fld_char_begin = etree.SubElement(begin_run, f'{{{NAMESPACES["w"]}}}fldChar')
        fld_char_begin.set(f'{{{NAMESPACES["w"]}}}fldCharType', 'begin')
        runs.append(begin_run)

        # Field code
        code_run = etree.Element(f'{{{NAMESPACES["w"]}}}r')
        instr_text = etree.SubElement(code_run, f'{{{NAMESPACES["w"]}}}instrText')
        instr_text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        instr_text.text = f' {field_def.field_code} '
        runs.append(code_run)

        # Field separator
        sep_run = etree.Element(f'{{{NAMESPACES["w"]}}}r')
        fld_char_sep = etree.SubElement(sep_run, f'{{{NAMESPACES["w"]}}}fldChar')
        fld_char_sep.set(f'{{{NAMESPACES["w"]}}}fldCharType', 'separate')
        runs.append(sep_run)

        # Display text
        if display_text:
            text_run = etree.Element(f'{{{NAMESPACES["w"]}}}r')
            t = etree.SubElement(text_run, f'{{{NAMESPACES["w"]}}}t')
            t.text = display_text
            runs.append(text_run)

        # Field end
        end_run = etree.Element(f'{{{NAMESPACES["w"]}}}r')
        fld_char_end = etree.SubElement(end_run, f'{{{NAMESPACES["w"]}}}fldChar')
        fld_char_end.set(f'{{{NAMESPACES["w"]}}}fldCharType', 'end')
        runs.append(end_run)

        return runs

    def build_bookmark_start(self, name: str, bookmark_id: Optional[str] = None) -> etree._Element:
        """Build a bookmarkStart element."""
        if bookmark_id is None:
            bookmark_id = str(self.next_bookmark_id)
            self.next_bookmark_id += 1

        elem = etree.Element(f'{{{NAMESPACES["w"]}}}bookmarkStart')
        elem.set(f'{{{NAMESPACES["w"]}}}id', bookmark_id)
        elem.set(f'{{{NAMESPACES["w"]}}}name', name)

        return elem

    def build_bookmark_end(self, bookmark_id: str) -> etree._Element:
        """Build a bookmarkEnd element."""
        elem = etree.Element(f'{{{NAMESPACES["w"]}}}bookmarkEnd')
        elem.set(f'{{{NAMESPACES["w"]}}}id', bookmark_id)

        return elem

    def extract_fields_from_paragraph(self, para_elem: etree._Element) -> list[dict]:
        """Extract field information from a paragraph element.

        Returns a list of field info dictionaries with their positions.
        """
        fields = []
        current_field = None
        field_start_index = None

        runs = para_elem.findall('w:r', namespaces=NAMESPACES)

        for i, run in enumerate(runs):
            fld_char = run.find('w:fldChar', namespaces=NAMESPACES)

            if fld_char is not None:
                fld_type = fld_char.get(f'{{{NAMESPACES["w"]}}}fldCharType')

                if fld_type == 'begin':
                    current_field = {
                        'start_index': i,
                        'field_code': '',
                        'display_runs': [],
                        'end_index': None
                    }
                    field_start_index = i

                elif fld_type == 'separate' and current_field:
                    current_field['separator_index'] = i

                elif fld_type == 'end' and current_field:
                    current_field['end_index'] = i
                    fields.append(current_field)
                    current_field = None

            elif current_field is not None:
                # This run is part of a field
                instr = run.find('w:instrText', namespaces=NAMESPACES)
                if instr is not None and instr.text:
                    current_field['field_code'] += instr.text

                # Check if we're past the separator (display text)
                if 'separator_index' in current_field and i > current_field['separator_index']:
                    current_field['display_runs'].append(run)

        return fields

    def preserve_fields_in_paragraph(self, old_para: etree._Element, new_para: etree._Element):
        """Copy field structures from old paragraph to new one.

        This ensures cross-references are preserved during rebuild.
        """
        # Extract fields from the old paragraph
        fields = self.extract_fields_from_paragraph(old_para)

        if not fields:
            return

        # For each field, we need to insert it at the appropriate position
        # in the new paragraph based on text position matching
        old_runs = old_para.findall('w:r', namespaces=NAMESPACES)
        new_runs = new_para.findall('w:r', namespaces=NAMESPACES)

        for field_info in fields:
            # Copy the entire field structure
            start_idx = field_info['start_index']
            end_idx = field_info['end_index']

            if start_idx < len(old_runs) and end_idx < len(old_runs):
                # Find insertion point in new paragraph
                # For simplicity, we'll append fields at the end of the new paragraph
                # In a more sophisticated implementation, we'd match text positions

                for i in range(start_idx, end_idx + 1):
                    run_copy = deepcopy(old_runs[i])
                    new_para.append(run_copy)

    def rebuild_bookmark(self, bookmark: Bookmark, content_elem: etree._Element) -> tuple[etree._Element, etree._Element]:
        """Create bookmark start and end elements for a piece of content.

        Returns (bookmarkStart, bookmarkEnd) tuple.
        """
        bookmark_id = bookmark.bookmark_id

        # Check if we need to remap the name
        new_name = bookmark.name
        if bookmark.name in self.bookmark_mappings:
            new_name = self.bookmark_mappings[bookmark.name].new_name

        start_elem = self.build_bookmark_start(new_name, bookmark_id)
        end_elem = self.build_bookmark_end(bookmark_id)

        return start_elem, end_elem

    def update_field_references(self, old_name: str, new_name: str):
        """Update all field references when a bookmark is renamed."""
        self.bookmark_mappings[old_name] = BookmarkMapping(
            original_name=old_name,
            new_name=new_name,
            bookmark_id=str(self.next_bookmark_id)
        )
        self.next_bookmark_id += 1

        # Update field definitions
        for field_def in self.field_definitions:
            if field_def.bookmark_name == old_name:
                field_def.bookmark_name = new_name
                # Rebuild the field code
                field_def.field_code = self._rebuild_field_code(field_def)

    def _rebuild_field_code(self, field_def: FieldDefinition) -> str:
        """Rebuild a field code string from a FieldDefinition."""
        parts = [field_def.field_type]

        if field_def.bookmark_name:
            parts.append(field_def.bookmark_name)

        parts.extend(field_def.switches)

        return ' '.join(parts)


def extract_hyperlinks(para_elem: etree._Element) -> list[dict]:
    """Extract hyperlink information from a paragraph.

    Hyperlinks in OOXML are represented as w:hyperlink elements.
    """
    hyperlinks = []

    for hyperlink in para_elem.findall('.//w:hyperlink', namespaces=NAMESPACES):
        rel_id = hyperlink.get(f'{{{NAMESPACES["r"]}}}id')
        anchor = hyperlink.get(f'{{{NAMESPACES["w"]}}}anchor')

        # Get display text
        text_parts = []
        for t in hyperlink.findall('.//w:t', namespaces=NAMESPACES):
            if t.text:
                text_parts.append(t.text)

        hyperlinks.append({
            'rel_id': rel_id,
            'anchor': anchor,
            'display_text': ''.join(text_parts),
            'element': hyperlink
        })

    return hyperlinks


def build_hyperlink_xml(rel_id: Optional[str] = None,
                        anchor: Optional[str] = None,
                        display_text: str = "",
                        formatting: dict = None) -> etree._Element:
    """Build a hyperlink element.

    Args:
        rel_id: Relationship ID for external hyperlinks
        anchor: Bookmark name for internal hyperlinks
        display_text: Text to display
        formatting: Optional formatting to apply to the text
    """
    hyperlink = etree.Element(f'{{{NAMESPACES["w"]}}}hyperlink')

    if rel_id:
        hyperlink.set(f'{{{NAMESPACES["r"]}}}id', rel_id)
    if anchor:
        hyperlink.set(f'{{{NAMESPACES["w"]}}}anchor', anchor)

    # Create run with text
    run = etree.SubElement(hyperlink, f'{{{NAMESPACES["w"]}}}r')

    # Add formatting (hyperlinks typically have underline and blue color)
    rPr = etree.SubElement(run, f'{{{NAMESPACES["w"]}}}rPr')

    # Default hyperlink style
    rStyle = etree.SubElement(rPr, f'{{{NAMESPACES["w"]}}}rStyle')
    rStyle.set(f'{{{NAMESPACES["w"]}}}val', 'Hyperlink')

    # Text
    t = etree.SubElement(run, f'{{{NAMESPACES["w"]}}}t')
    t.text = display_text

    return hyperlink
