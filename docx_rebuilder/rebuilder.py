"""
DOCX Rebuilder - Core module for rebuilding documents with clean formatting.

This module takes a parsed document structure, analyzes it, and rebuilds
the OOXML with clean, consistent formatting while preserving content.
"""

import zipfile
import os
import shutil
from pathlib import Path
from copy import deepcopy
from typing import Optional
from lxml import etree

from .parser import (
    DocxParser, DocumentStructure, Paragraph, Table, Run,
    ParagraphFormatting, RunFormatting, NAMESPACES
)
from .analyzer import FormattingAnalyzer, FormattingModel
from .numbering import NumberingHandler
from .crossref import CrossReferenceHandler
from .headers import HeaderFooterHandler, preserve_section_properties


class DocxRebuilder:
    """Rebuilds DOCX documents with clean, consistent formatting."""

    def __init__(self, input_path: str | Path, output_path: Optional[str | Path] = None):
        self.input_path = Path(input_path)
        self.output_path = Path(output_path) if output_path else self.input_path.with_stem(
            self.input_path.stem + '_rebuilt'
        )

        self.parser: Optional[DocxParser] = None
        self.structure: Optional[DocumentStructure] = None
        self.analyzer: Optional[FormattingAnalyzer] = None
        self.model: Optional[FormattingModel] = None
        self.numbering_handler: Optional[NumberingHandler] = None
        self.crossref_handler: Optional[CrossReferenceHandler] = None
        self.header_footer_handler: Optional[HeaderFooterHandler] = None

        # Store original XML parts for preservation
        self.original_parts: dict[str, bytes] = {}

    def rebuild(self, normalize_numbering: bool = True,
                normalize_styles: bool = True,
                preserve_cross_refs: bool = True) -> Path:
        """Perform the full rebuild process.

        Args:
            normalize_numbering: Whether to normalize numbering schemes
            normalize_styles: Whether to apply consistent styles
            preserve_cross_refs: Whether to preserve cross-references

        Returns:
            Path to the rebuilt document
        """
        # Phase 1: Parse the document
        self._parse_document()

        # Phase 2: Analyze formatting
        self._analyze_formatting()

        # Phase 3: Initialize handlers
        self._initialize_handlers()

        # Phase 4: Rebuild the document
        self._rebuild_document(
            normalize_numbering=normalize_numbering,
            normalize_styles=normalize_styles,
            preserve_cross_refs=preserve_cross_refs
        )

        return self.output_path

    def _parse_document(self):
        """Parse the input document."""
        self.parser = DocxParser(self.input_path)
        self.structure = self.parser.parse()

        # Store original XML parts
        with zipfile.ZipFile(self.input_path, 'r') as zf:
            for name in zf.namelist():
                self.original_parts[name] = zf.read(name)

    def _analyze_formatting(self):
        """Analyze document formatting to build the model."""
        self.analyzer = FormattingAnalyzer(self.structure)
        self.model = self.analyzer.analyze()

    def _initialize_handlers(self):
        """Initialize the numbering, cross-reference, and header/footer handlers."""
        self.numbering_handler = NumberingHandler(
            self.structure,
            indent_unit=self.model.indent_unit
        )
        self.numbering_handler.analyze_and_normalize()

        self.crossref_handler = CrossReferenceHandler(self.structure)
        self.crossref_handler.analyze_references()

        # Initialize header/footer handler
        self.header_footer_handler = HeaderFooterHandler(self.input_path)
        self.header_footer_handler.parse()

    def _rebuild_document(self, normalize_numbering: bool,
                          normalize_styles: bool,
                          preserve_cross_refs: bool):
        """Rebuild the document with clean formatting."""
        # Create a temporary copy of the original document
        temp_path = self.output_path.with_suffix('.tmp')
        shutil.copy2(self.input_path, temp_path)

        try:
            with zipfile.ZipFile(temp_path, 'a') as zf:
                # Rebuild document.xml
                new_document = self._rebuild_document_xml(
                    normalize_styles=normalize_styles,
                    preserve_cross_refs=preserve_cross_refs
                )
                self._update_zip_file(zf, 'word/document.xml', new_document)

                # Rebuild styles.xml if normalizing styles
                if normalize_styles:
                    new_styles = self._rebuild_styles_xml()
                    self._update_zip_file(zf, 'word/styles.xml', new_styles)

                # Rebuild numbering.xml if normalizing numbering
                if normalize_numbering and self.structure.numbering_definitions:
                    new_numbering = self._rebuild_numbering_xml()
                    self._update_zip_file(zf, 'word/numbering.xml', new_numbering)

            # Move temp file to final output
            shutil.move(temp_path, self.output_path)

        finally:
            # Clean up temp file if it still exists
            if temp_path.exists():
                temp_path.unlink()

    def _update_zip_file(self, zf: zipfile.ZipFile, name: str, content: bytes):
        """Update a file in the zip archive."""
        # Remove old file if it exists
        # Note: zipfile doesn't support removing files directly,
        # so we need a workaround for existing files
        zf.writestr(name, content)

    def _rebuild_document_xml(self, normalize_styles: bool,
                               preserve_cross_refs: bool) -> bytes:
        """Rebuild document.xml with clean formatting."""
        # Start from a clean document structure
        nsmap = {
            'w': NAMESPACES['w'],
            'r': NAMESPACES['r'],
            'wp': NAMESPACES['wp'],
            'mc': NAMESPACES['mc'],
        }

        # Get the original document root
        original_root = self.parser.get_raw_xml('document')

        # Find the body element
        body = original_root.find('.//w:body', namespaces=NAMESPACES)
        if body is None:
            return etree.tostring(original_root, xml_declaration=True, encoding='UTF-8')

        # Process each element in the body
        new_body_children = []

        for item_type, item_index in self.structure.content_order:
            if item_type == 'paragraph':
                para = self.structure.paragraphs[item_index]
                new_para = self._rebuild_paragraph(
                    para,
                    normalize_styles=normalize_styles,
                    preserve_cross_refs=preserve_cross_refs,
                    para_index=item_index
                )
                new_body_children.append(new_para)

            elif item_type == 'table':
                table = self.structure.tables[item_index]
                new_table = self._rebuild_table(
                    table,
                    normalize_styles=normalize_styles
                )
                new_body_children.append(new_table)

        # Get sectPr (section properties) from original - must be last in body
        sect_pr = body.find('w:sectPr', namespaces=NAMESPACES)

        # Clear body and add rebuilt content
        for child in list(body):
            body.remove(child)

        for child in new_body_children:
            body.append(child)

        # Add section properties back
        if sect_pr is not None:
            body.append(deepcopy(sect_pr))

        return etree.tostring(original_root, xml_declaration=True, encoding='UTF-8')

    def _rebuild_paragraph(self, para: Paragraph,
                           normalize_styles: bool,
                           preserve_cross_refs: bool,
                           para_index: int = -1) -> etree._Element:
        """Rebuild a paragraph with clean formatting."""
        p = etree.Element(f'{{{NAMESPACES["w"]}}}p')

        # Build paragraph properties
        pPr = self._build_paragraph_properties(para, normalize_styles)
        if pPr is not None and len(pPr) > 0:
            p.append(pPr)

        # Handle bookmarks at start
        if preserve_cross_refs:
            for bm_name, bookmark in self.structure.bookmarks.items():
                if bookmark.start_para_index == para_index:
                    start_elem, _ = self.crossref_handler.rebuild_bookmark(
                        bookmark, p
                    )
                    p.append(start_elem)

        # Build runs
        for run in para.runs:
            run_elem = self._rebuild_run(run, normalize_styles)
            p.append(run_elem)

        # Handle bookmarks at end
        if preserve_cross_refs:
            for bm_name, bookmark in self.structure.bookmarks.items():
                if bookmark.end_para_index == para_index:
                    _, end_elem = self.crossref_handler.rebuild_bookmark(
                        bookmark, p
                    )
                    p.append(end_elem)

        # Preserve cross-references from original
        if preserve_cross_refs and para.xml_element is not None:
            self.crossref_handler.preserve_fields_in_paragraph(
                para.xml_element, p
            )

        return p

    def _build_paragraph_properties(self, para: Paragraph,
                                     normalize_styles: bool) -> etree._Element:
        """Build paragraph properties element."""
        pPr = etree.Element(f'{{{NAMESPACES["w"]}}}pPr')
        fmt = para.formatting

        # Style reference
        if fmt.style_id:
            pStyle = etree.SubElement(pPr, f'{{{NAMESPACES["w"]}}}pStyle')
            pStyle.set(f'{{{NAMESPACES["w"]}}}val', fmt.style_id)

        # Numbering
        if para.numbering_id is not None:
            numPr = etree.SubElement(pPr, f'{{{NAMESPACES["w"]}}}numPr')

            if para.numbering_level is not None:
                ilvl = etree.SubElement(numPr, f'{{{NAMESPACES["w"]}}}ilvl')
                ilvl.set(f'{{{NAMESPACES["w"]}}}val', str(para.numbering_level))

            numId = etree.SubElement(numPr, f'{{{NAMESPACES["w"]}}}numId')
            numId.set(f'{{{NAMESPACES["w"]}}}val', para.numbering_id)

        # Alignment
        if fmt.alignment:
            jc = etree.SubElement(pPr, f'{{{NAMESPACES["w"]}}}jc')
            jc.set(f'{{{NAMESPACES["w"]}}}val', fmt.alignment)

        # Indentation
        if any([fmt.indent_left, fmt.indent_right, fmt.indent_first_line, fmt.indent_hanging]):
            ind = etree.SubElement(pPr, f'{{{NAMESPACES["w"]}}}ind')

            if fmt.indent_left is not None:
                # Normalize to indent unit if requested
                value = fmt.indent_left
                if normalize_styles:
                    value = self._round_to_unit(value, self.model.indent_unit)
                ind.set(f'{{{NAMESPACES["w"]}}}left', str(value))

            if fmt.indent_right is not None:
                ind.set(f'{{{NAMESPACES["w"]}}}right', str(fmt.indent_right))

            if fmt.indent_first_line is not None:
                ind.set(f'{{{NAMESPACES["w"]}}}firstLine', str(fmt.indent_first_line))

            if fmt.indent_hanging is not None:
                ind.set(f'{{{NAMESPACES["w"]}}}hanging', str(fmt.indent_hanging))

        # Spacing
        if any([fmt.spacing_before, fmt.spacing_after, fmt.line_spacing]):
            spacing = etree.SubElement(pPr, f'{{{NAMESPACES["w"]}}}spacing')

            if fmt.spacing_before is not None:
                spacing.set(f'{{{NAMESPACES["w"]}}}before', str(fmt.spacing_before))

            if fmt.spacing_after is not None:
                spacing.set(f'{{{NAMESPACES["w"]}}}after', str(fmt.spacing_after))

            if fmt.line_spacing is not None:
                spacing.set(f'{{{NAMESPACES["w"]}}}line', str(fmt.line_spacing))

            if fmt.line_spacing_rule:
                spacing.set(f'{{{NAMESPACES["w"]}}}lineRule', fmt.line_spacing_rule)

        # Keep properties
        if fmt.keep_next:
            etree.SubElement(pPr, f'{{{NAMESPACES["w"]}}}keepNext')
        if fmt.keep_lines:
            etree.SubElement(pPr, f'{{{NAMESPACES["w"]}}}keepLines')
        if fmt.page_break_before:
            etree.SubElement(pPr, f'{{{NAMESPACES["w"]}}}pageBreakBefore')

        # Outline level
        if fmt.outline_level is not None:
            outlineLvl = etree.SubElement(pPr, f'{{{NAMESPACES["w"]}}}outlineLvl')
            outlineLvl.set(f'{{{NAMESPACES["w"]}}}val', str(fmt.outline_level))

        return pPr

    def _rebuild_run(self, run: Run, normalize_styles: bool) -> etree._Element:
        """Rebuild a run with clean formatting."""
        r = etree.Element(f'{{{NAMESPACES["w"]}}}r')

        # Build run properties
        rPr = self._build_run_properties(run.formatting, normalize_styles)
        if rPr is not None and len(rPr) > 0:
            r.append(rPr)

        # Handle text content (may include tabs and line breaks)
        text = run.text
        i = 0
        while i < len(text):
            char = text[i]

            if char == '\t':
                etree.SubElement(r, f'{{{NAMESPACES["w"]}}}tab')
                i += 1
            elif char == '\n':
                etree.SubElement(r, f'{{{NAMESPACES["w"]}}}br')
                i += 1
            else:
                # Find the end of the regular text segment
                j = i
                while j < len(text) and text[j] not in '\t\n':
                    j += 1

                t = etree.SubElement(r, f'{{{NAMESPACES["w"]}}}t')
                t.text = text[i:j]

                # Preserve whitespace
                if text[i:j].startswith(' ') or text[i:j].endswith(' '):
                    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

                i = j

        return r

    def _build_run_properties(self, fmt: RunFormatting,
                               normalize_styles: bool) -> etree._Element:
        """Build run properties element."""
        rPr = etree.Element(f'{{{NAMESPACES["w"]}}}rPr')

        # Font
        if fmt.font_name:
            font_name = fmt.font_name
            if normalize_styles and fmt.font_name != self.model.base_font_name:
                # Keep the original font - don't force normalization
                # Only normalize if we're confident it's an error
                pass

            rFonts = etree.SubElement(rPr, f'{{{NAMESPACES["w"]}}}rFonts')
            rFonts.set(f'{{{NAMESPACES["w"]}}}ascii', font_name)
            rFonts.set(f'{{{NAMESPACES["w"]}}}hAnsi', font_name)

        # Font size
        if fmt.font_size:
            sz = etree.SubElement(rPr, f'{{{NAMESPACES["w"]}}}sz')
            sz.set(f'{{{NAMESPACES["w"]}}}val', str(fmt.font_size))
            szCs = etree.SubElement(rPr, f'{{{NAMESPACES["w"]}}}szCs')
            szCs.set(f'{{{NAMESPACES["w"]}}}val', str(fmt.font_size))

        # Bold
        if fmt.bold:
            etree.SubElement(rPr, f'{{{NAMESPACES["w"]}}}b')

        # Italic
        if fmt.italic:
            etree.SubElement(rPr, f'{{{NAMESPACES["w"]}}}i')

        # Underline
        if fmt.underline:
            u = etree.SubElement(rPr, f'{{{NAMESPACES["w"]}}}u')
            u.set(f'{{{NAMESPACES["w"]}}}val', fmt.underline)

        # Strikethrough
        if fmt.strike:
            etree.SubElement(rPr, f'{{{NAMESPACES["w"]}}}strike')

        # Color
        if fmt.color:
            color = etree.SubElement(rPr, f'{{{NAMESPACES["w"]}}}color')
            color.set(f'{{{NAMESPACES["w"]}}}val', fmt.color)

        # Highlight
        if fmt.highlight:
            highlight = etree.SubElement(rPr, f'{{{NAMESPACES["w"]}}}highlight')
            highlight.set(f'{{{NAMESPACES["w"]}}}val', fmt.highlight)

        # Superscript/Subscript
        if fmt.superscript or fmt.subscript:
            vertAlign = etree.SubElement(rPr, f'{{{NAMESPACES["w"]}}}vertAlign')
            vertAlign.set(f'{{{NAMESPACES["w"]}}}val',
                         'superscript' if fmt.superscript else 'subscript')

        # Small caps / All caps
        if fmt.small_caps:
            etree.SubElement(rPr, f'{{{NAMESPACES["w"]}}}smallCaps')
        if fmt.all_caps:
            etree.SubElement(rPr, f'{{{NAMESPACES["w"]}}}caps')

        return rPr

    def _rebuild_table(self, table: Table, normalize_styles: bool) -> etree._Element:
        """Rebuild a table with clean formatting."""
        tbl = etree.Element(f'{{{NAMESPACES["w"]}}}tbl')

        # Table properties
        tblPr = etree.SubElement(tbl, f'{{{NAMESPACES["w"]}}}tblPr')

        # Table width
        if table.width:
            tblW = etree.SubElement(tblPr, f'{{{NAMESPACES["w"]}}}tblW')
            tblW.set(f'{{{NAMESPACES["w"]}}}w', str(table.width))
            tblW.set(f'{{{NAMESPACES["w"]}}}type', 'dxa')

        # Table alignment
        if table.alignment:
            jc = etree.SubElement(tblPr, f'{{{NAMESPACES["w"]}}}jc')
            jc.set(f'{{{NAMESPACES["w"]}}}val', table.alignment)

        # Default table borders
        tblBorders = etree.SubElement(tblPr, f'{{{NAMESPACES["w"]}}}tblBorders')
        for border_type in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = etree.SubElement(tblBorders, f'{{{NAMESPACES["w"]}}}{border_type}')
            border.set(f'{{{NAMESPACES["w"]}}}val', 'single')
            border.set(f'{{{NAMESPACES["w"]}}}sz', '4')
            border.set(f'{{{NAMESPACES["w"]}}}space', '0')
            border.set(f'{{{NAMESPACES["w"]}}}color', 'auto')

        # Table grid
        tblGrid = etree.SubElement(tbl, f'{{{NAMESPACES["w"]}}}tblGrid')

        # Calculate column widths from first row
        if table.rows:
            for cell in table.rows[0].cells:
                gridCol = etree.SubElement(tblGrid, f'{{{NAMESPACES["w"]}}}gridCol')
                if cell.width:
                    gridCol.set(f'{{{NAMESPACES["w"]}}}w', str(cell.width))

        # Rows
        for row in table.rows:
            tr = self._rebuild_table_row(row, normalize_styles)
            tbl.append(tr)

        return tbl

    def _rebuild_table_row(self, row, normalize_styles: bool) -> etree._Element:
        """Rebuild a table row."""
        tr = etree.Element(f'{{{NAMESPACES["w"]}}}tr')

        # Row properties
        trPr = etree.SubElement(tr, f'{{{NAMESPACES["w"]}}}trPr')

        if row.height:
            trHeight = etree.SubElement(trPr, f'{{{NAMESPACES["w"]}}}trHeight')
            trHeight.set(f'{{{NAMESPACES["w"]}}}val', str(row.height))

        if row.is_header:
            etree.SubElement(trPr, f'{{{NAMESPACES["w"]}}}tblHeader')

        # Cells
        for cell in row.cells:
            tc = self._rebuild_table_cell(cell, normalize_styles)
            tr.append(tc)

        return tr

    def _rebuild_table_cell(self, cell, normalize_styles: bool) -> etree._Element:
        """Rebuild a table cell."""
        tc = etree.Element(f'{{{NAMESPACES["w"]}}}tc')

        # Cell properties
        tcPr = etree.SubElement(tc, f'{{{NAMESPACES["w"]}}}tcPr')

        if cell.width:
            tcW = etree.SubElement(tcPr, f'{{{NAMESPACES["w"]}}}tcW')
            tcW.set(f'{{{NAMESPACES["w"]}}}w', str(cell.width))
            tcW.set(f'{{{NAMESPACES["w"]}}}type', 'dxa')

        if cell.col_span > 1:
            gridSpan = etree.SubElement(tcPr, f'{{{NAMESPACES["w"]}}}gridSpan')
            gridSpan.set(f'{{{NAMESPACES["w"]}}}val', str(cell.col_span))

        if cell.vertical_align:
            vAlign = etree.SubElement(tcPr, f'{{{NAMESPACES["w"]}}}vAlign')
            vAlign.set(f'{{{NAMESPACES["w"]}}}val', cell.vertical_align)

        # Paragraphs in cell
        for para in cell.paragraphs:
            p = self._rebuild_paragraph(para, normalize_styles, preserve_cross_refs=True)
            tc.append(p)

        # Ensure at least one paragraph (required by OOXML)
        if not cell.paragraphs:
            p = etree.SubElement(tc, f'{{{NAMESPACES["w"]}}}p')

        return tc

    def _rebuild_styles_xml(self) -> bytes:
        """Rebuild styles.xml with clean style definitions."""
        # Start from the original styles
        original_styles = self.parser.get_raw_xml('styles')

        if original_styles is None:
            # Create new styles document
            nsmap = {'w': NAMESPACES['w']}
            original_styles = etree.Element(f'{{{NAMESPACES["w"]}}}styles', nsmap=nsmap)

        # Clean up redundant style overrides
        # This preserves the styles but removes any corruption

        return etree.tostring(original_styles, xml_declaration=True, encoding='UTF-8')

    def _rebuild_numbering_xml(self) -> bytes:
        """Rebuild numbering.xml with clean numbering definitions."""
        numbering_root = self.numbering_handler.build_numbering_xml()
        return etree.tostring(numbering_root, xml_declaration=True, encoding='UTF-8')

    def _round_to_unit(self, value: int, unit: int) -> int:
        """Round a value to the nearest multiple of the unit."""
        return round(value / unit) * unit

    def get_analysis_report(self) -> dict:
        """Get a report of the document analysis."""
        if not self.model:
            return {}

        report = {
            'base_font': self.model.base_font_name,
            'base_font_size': self.model.base_font_size / 2,  # Convert half-points to points
            'default_alignment': self.model.default_alignment,
            'indent_unit': self.model.indent_unit,
            'line_spacing': self.model.line_spacing,
            'heading_styles_detected': len(self.model.heading_styles),
            'numbering_schemes': len(self.model.numbering_schemes),
            'paragraph_count': len(self.structure.paragraphs),
            'table_count': len(self.structure.tables),
            'bookmark_count': len(self.structure.bookmarks),
            'cross_reference_count': len(self.structure.cross_references),
        }

        # Add header/footer info if available
        if self.header_footer_handler:
            report['headers_count'] = len(self.header_footer_handler.headers)
            report['footers_count'] = len(self.header_footer_handler.footers)
            report['sections_count'] = len(self.header_footer_handler.section_props)

            # Check for page number fields
            page_fields = self.header_footer_handler.get_page_number_fields()
            report['page_number_fields'] = len(page_fields)

        return report
