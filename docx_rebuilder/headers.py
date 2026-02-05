"""
Header/Footer Handler - Preserves and processes document headers and footers.

This module handles the preservation of headers, footers, page numbers,
and other section-level content in Word documents.
"""

import zipfile
from dataclasses import dataclass, field
from typing import Optional
from pathlib import Path
from copy import deepcopy
from lxml import etree

from .parser import NAMESPACES, Paragraph, Table


@dataclass
class HeaderFooter:
    """Represents a header or footer in the document."""
    hf_type: str  # 'header' or 'footer'
    hf_id: str  # e.g., 'header1', 'footer1'
    rel_id: str  # Relationship ID (e.g., 'rId8')
    xml_content: bytes
    xml_tree: Optional[etree._Element] = None
    position: str = 'default'  # 'default', 'first', 'even'


@dataclass
class SectionProperties:
    """Section properties including page setup and header/footer refs."""
    page_width: Optional[int] = None
    page_height: Optional[int] = None
    margin_top: Optional[int] = None
    margin_bottom: Optional[int] = None
    margin_left: Optional[int] = None
    margin_right: Optional[int] = None
    header_distance: Optional[int] = None
    footer_distance: Optional[int] = None
    page_number_start: Optional[int] = None
    headers: dict[str, str] = field(default_factory=dict)  # position -> rel_id
    footers: dict[str, str] = field(default_factory=dict)  # position -> rel_id
    xml_element: Optional[etree._Element] = None


class HeaderFooterHandler:
    """Handles header and footer preservation and processing."""

    def __init__(self, docx_path: Path):
        self.docx_path = docx_path
        self.headers: dict[str, HeaderFooter] = {}
        self.footers: dict[str, HeaderFooter] = {}
        self.relationships: dict[str, dict] = {}
        self.section_props: list[SectionProperties] = []

    def parse(self):
        """Parse headers, footers, and section properties from the document."""
        with zipfile.ZipFile(self.docx_path, 'r') as zf:
            # Parse document relationships
            if 'word/_rels/document.xml.rels' in zf.namelist():
                self._parse_relationships(zf)

            # Parse headers and footers
            for name in zf.namelist():
                if name.startswith('word/header') and name.endswith('.xml'):
                    self._parse_header_footer(zf, name, 'header')
                elif name.startswith('word/footer') and name.endswith('.xml'):
                    self._parse_header_footer(zf, name, 'footer')

            # Parse section properties from document.xml
            if 'word/document.xml' in zf.namelist():
                self._parse_section_properties(zf)

    def _parse_relationships(self, zf: zipfile.ZipFile):
        """Parse document relationships to find header/footer references."""
        rels_content = zf.read('word/_rels/document.xml.rels')
        rels_tree = etree.fromstring(rels_content)

        for rel in rels_tree.iter():
            if rel.tag.endswith('Relationship'):
                rel_id = rel.get('Id')
                rel_type = rel.get('Type', '')
                target = rel.get('Target', '')

                self.relationships[rel_id] = {
                    'type': rel_type,
                    'target': target,
                    'is_header': 'header' in rel_type.lower(),
                    'is_footer': 'footer' in rel_type.lower(),
                }

    def _parse_header_footer(self, zf: zipfile.ZipFile, path: str, hf_type: str):
        """Parse a header or footer XML file."""
        content = zf.read(path)
        tree = etree.fromstring(content)

        # Extract the ID from the filename (e.g., 'header1' from 'word/header1.xml')
        hf_id = path.replace('word/', '').replace('.xml', '')

        # Find the relationship ID for this header/footer
        rel_id = None
        for rid, rel_info in self.relationships.items():
            if rel_info['target'] == f'{hf_id}.xml':
                rel_id = rid
                break

        hf = HeaderFooter(
            hf_type=hf_type,
            hf_id=hf_id,
            rel_id=rel_id or '',
            xml_content=content,
            xml_tree=tree,
        )

        if hf_type == 'header':
            self.headers[hf_id] = hf
        else:
            self.footers[hf_id] = hf

    def _parse_section_properties(self, zf: zipfile.ZipFile):
        """Parse section properties from document.xml."""
        doc_content = zf.read('word/document.xml')
        doc_tree = etree.fromstring(doc_content)

        # Find all sectPr elements (in body and in paragraphs)
        sect_prs = doc_tree.xpath('//w:sectPr', namespaces=NAMESPACES)

        for sect_pr in sect_prs:
            props = SectionProperties(xml_element=deepcopy(sect_pr))

            # Page size
            pg_sz = sect_pr.find('w:pgSz', namespaces=NAMESPACES)
            if pg_sz is not None:
                w = pg_sz.get(f'{{{NAMESPACES["w"]}}}w')
                h = pg_sz.get(f'{{{NAMESPACES["w"]}}}h')
                props.page_width = int(w) if w else None
                props.page_height = int(h) if h else None

            # Page margins
            pg_mar = sect_pr.find('w:pgMar', namespaces=NAMESPACES)
            if pg_mar is not None:
                for attr, prop in [('top', 'margin_top'), ('bottom', 'margin_bottom'),
                                   ('left', 'margin_left'), ('right', 'margin_right'),
                                   ('header', 'header_distance'), ('footer', 'footer_distance')]:
                    val = pg_mar.get(f'{{{NAMESPACES["w"]}}}{attr}')
                    if val:
                        setattr(props, prop, int(val))

            # Page number start
            pg_num_type = sect_pr.find('w:pgNumType', namespaces=NAMESPACES)
            if pg_num_type is not None:
                start = pg_num_type.get(f'{{{NAMESPACES["w"]}}}start')
                if start:
                    props.page_number_start = int(start)

            # Header references
            for header_ref in sect_pr.findall('w:headerReference', namespaces=NAMESPACES):
                ref_type = header_ref.get(f'{{{NAMESPACES["w"]}}}type', 'default')
                rel_id = header_ref.get(f'{{{NAMESPACES["r"]}}}id')
                if rel_id:
                    props.headers[ref_type] = rel_id

            # Footer references
            for footer_ref in sect_pr.findall('w:footerReference', namespaces=NAMESPACES):
                ref_type = footer_ref.get(f'{{{NAMESPACES["w"]}}}type', 'default')
                rel_id = footer_ref.get(f'{{{NAMESPACES["r"]}}}id')
                if rel_id:
                    props.footers[ref_type] = rel_id

            self.section_props.append(props)

    def get_page_number_fields(self) -> list[dict]:
        """Find all page number fields in headers and footers."""
        fields = []

        for hf_id, hf in {**self.headers, **self.footers}.items():
            if hf.xml_tree is None:
                continue

            # Look for PAGE fields
            for instr in hf.xml_tree.xpath('.//w:instrText', namespaces=NAMESPACES):
                text = instr.text or ''
                if 'PAGE' in text.upper() or 'NUMPAGES' in text.upper():
                    fields.append({
                        'location': hf_id,
                        'type': hf.hf_type,
                        'field_code': text.strip(),
                    })

            # Also check for simple field codes
            for fld_simple in hf.xml_tree.xpath('.//w:fldSimple', namespaces=NAMESPACES):
                instr = fld_simple.get(f'{{{NAMESPACES["w"]}}}instr', '')
                if 'PAGE' in instr.upper() or 'NUMPAGES' in instr.upper():
                    fields.append({
                        'location': hf_id,
                        'type': hf.hf_type,
                        'field_code': instr.strip(),
                    })

        return fields

    def rebuild_header_footer(self, hf: HeaderFooter, normalize_styles: bool = True) -> bytes:
        """Rebuild a header or footer with clean formatting.

        For now, we preserve headers/footers as-is since they typically
        contain page numbers and other dynamic content that should not
        be modified.
        """
        # Return original content - headers/footers are preserved as-is
        # to maintain page numbering, dates, and other dynamic fields
        return hf.xml_content

    def get_all_header_footer_files(self) -> dict[str, bytes]:
        """Get all header and footer files for inclusion in rebuilt document."""
        files = {}

        for hf_id, hf in self.headers.items():
            files[f'word/{hf_id}.xml'] = self.rebuild_header_footer(hf)

        for hf_id, hf in self.footers.items():
            files[f'word/{hf_id}.xml'] = self.rebuild_header_footer(hf)

        return files


def extract_page_fields(xml_element: etree._Element) -> list[dict]:
    """Extract page-related fields from an XML element.

    Finds PAGE, NUMPAGES, SECTIONPAGES, and related fields.
    """
    fields = []

    # Complex fields (fldChar + instrText)
    field_codes = xml_element.xpath('.//w:instrText', namespaces=NAMESPACES)
    for instr in field_codes:
        text = (instr.text or '').strip().upper()
        if any(f in text for f in ['PAGE', 'NUMPAGES', 'SECTIONPAGES', 'DATE', 'TIME']):
            fields.append({
                'type': 'complex',
                'code': instr.text.strip(),
                'element': instr,
            })

    # Simple fields
    for fld_simple in xml_element.xpath('.//w:fldSimple', namespaces=NAMESPACES):
        instr = fld_simple.get(f'{{{NAMESPACES["w"]}}}instr', '').strip().upper()
        if any(f in instr for f in ['PAGE', 'NUMPAGES', 'SECTIONPAGES', 'DATE', 'TIME']):
            fields.append({
                'type': 'simple',
                'code': fld_simple.get(f'{{{NAMESPACES["w"]}}}instr', '').strip(),
                'element': fld_simple,
            })

    return fields


def preserve_section_properties(original_sect_pr: etree._Element) -> etree._Element:
    """Create a copy of section properties, preserving all settings.

    This ensures page setup, headers, footers, and page numbering
    are all maintained in the rebuilt document.
    """
    return deepcopy(original_sect_pr)
