"""
DOCX Parser - Extracts document structure and formatting from Word documents.
"""

import zipfile
import copy
from dataclasses import dataclass, field
from typing import Optional
from pathlib import Path
from lxml import etree

# OOXML namespaces
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}


@dataclass
class RunFormatting:
    """Character-level formatting for a run of text."""
    bold: bool = False
    italic: bool = False
    underline: Optional[str] = None  # single, double, etc.
    strike: bool = False
    font_name: Optional[str] = None
    font_size: Optional[int] = None  # in half-points
    color: Optional[str] = None
    highlight: Optional[str] = None
    superscript: bool = False
    subscript: bool = False
    small_caps: bool = False
    all_caps: bool = False


@dataclass
class ParagraphFormatting:
    """Paragraph-level formatting."""
    alignment: Optional[str] = None  # left, center, right, both (justify)
    indent_left: Optional[int] = None  # in twips
    indent_right: Optional[int] = None
    indent_first_line: Optional[int] = None
    indent_hanging: Optional[int] = None
    spacing_before: Optional[int] = None
    spacing_after: Optional[int] = None
    line_spacing: Optional[int] = None
    line_spacing_rule: Optional[str] = None
    outline_level: Optional[int] = None
    keep_next: bool = False
    keep_lines: bool = False
    page_break_before: bool = False
    style_id: Optional[str] = None


@dataclass
class Run:
    """A run of text with consistent formatting."""
    text: str
    formatting: RunFormatting
    xml_element: object = None  # Original XML element


@dataclass
class Paragraph:
    """A paragraph with its content and formatting."""
    runs: list[Run] = field(default_factory=list)
    formatting: ParagraphFormatting = field(default_factory=ParagraphFormatting)
    numbering_id: Optional[str] = None
    numbering_level: Optional[int] = None
    xml_element: object = None

    @property
    def text(self) -> str:
        return ''.join(run.text for run in self.runs)


@dataclass
class TableCell:
    """A table cell containing paragraphs."""
    paragraphs: list[Paragraph] = field(default_factory=list)
    width: Optional[int] = None
    col_span: int = 1
    row_span: int = 1
    vertical_align: Optional[str] = None
    xml_element: object = None


@dataclass
class TableRow:
    """A table row containing cells."""
    cells: list[TableCell] = field(default_factory=list)
    height: Optional[int] = None
    is_header: bool = False
    xml_element: object = None


@dataclass
class Table:
    """A table with rows and cells."""
    rows: list[TableRow] = field(default_factory=list)
    width: Optional[int] = None
    alignment: Optional[str] = None
    xml_element: object = None


@dataclass
class CrossReference:
    """A cross-reference to another part of the document."""
    field_code: str
    bookmark_name: Optional[str] = None
    ref_type: str = "unknown"  # REF, PAGEREF, SEQ, etc.
    display_text: str = ""
    xml_element: object = None


@dataclass
class Bookmark:
    """A bookmark in the document."""
    name: str
    bookmark_id: str
    start_para_index: int
    end_para_index: int


@dataclass
class NumberingDefinition:
    """A numbering definition from numbering.xml."""
    num_id: str
    abstract_num_id: str
    levels: dict = field(default_factory=dict)  # level -> NumberingLevel


@dataclass
class NumberingLevel:
    """A single level in a numbering definition."""
    level: int
    start: int = 1
    num_fmt: str = "decimal"  # decimal, lowerLetter, upperLetter, lowerRoman, upperRoman, bullet
    level_text: str = "%1."
    indent_left: Optional[int] = None
    indent_hanging: Optional[int] = None
    alignment: str = "left"


@dataclass
class Style:
    """A document style definition."""
    style_id: str
    name: str
    style_type: str  # paragraph, character, numbering, table
    based_on: Optional[str] = None
    paragraph_formatting: Optional[ParagraphFormatting] = None
    run_formatting: Optional[RunFormatting] = None
    xml_element: object = None


@dataclass
class DocumentStructure:
    """Complete document structure."""
    paragraphs: list[Paragraph] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)
    styles: dict[str, Style] = field(default_factory=dict)
    numbering_definitions: dict[str, NumberingDefinition] = field(default_factory=dict)
    bookmarks: dict[str, Bookmark] = field(default_factory=dict)
    cross_references: list[CrossReference] = field(default_factory=list)
    # Track content order (paragraph index or table reference)
    content_order: list = field(default_factory=list)


class DocxParser:
    """Parses DOCX files and extracts document structure."""

    def __init__(self, file_path: str | Path):
        self.file_path = Path(file_path)
        self.zip_file = None
        self.document_xml = None
        self.styles_xml = None
        self.numbering_xml = None
        self.relationships = {}

    def parse(self) -> DocumentStructure:
        """Parse the DOCX file and return the document structure."""
        structure = DocumentStructure()

        with zipfile.ZipFile(self.file_path, 'r') as zf:
            self.zip_file = zf

            # Parse main document
            if 'word/document.xml' in zf.namelist():
                self.document_xml = etree.parse(zf.open('word/document.xml'))

            # Parse styles
            if 'word/styles.xml' in zf.namelist():
                self.styles_xml = etree.parse(zf.open('word/styles.xml'))
                structure.styles = self._parse_styles()

            # Parse numbering
            if 'word/numbering.xml' in zf.namelist():
                self.numbering_xml = etree.parse(zf.open('word/numbering.xml'))
                structure.numbering_definitions = self._parse_numbering()

            # Parse relationships
            if 'word/_rels/document.xml.rels' in zf.namelist():
                self.relationships = self._parse_relationships(zf)

            # Parse document body
            if self.document_xml is not None:
                self._parse_body(structure)

        return structure

    def _parse_styles(self) -> dict[str, Style]:
        """Parse styles.xml and return style definitions."""
        styles = {}

        for style_elem in self.styles_xml.xpath('//w:style', namespaces=NAMESPACES):
            style_id = style_elem.get(f'{{{NAMESPACES["w"]}}}styleId')
            style_type = style_elem.get(f'{{{NAMESPACES["w"]}}}type')

            name_elem = style_elem.find('w:name', namespaces=NAMESPACES)
            name = name_elem.get(f'{{{NAMESPACES["w"]}}}val') if name_elem is not None else style_id

            based_on_elem = style_elem.find('w:basedOn', namespaces=NAMESPACES)
            based_on = based_on_elem.get(f'{{{NAMESPACES["w"]}}}val') if based_on_elem is not None else None

            style = Style(
                style_id=style_id,
                name=name,
                style_type=style_type,
                based_on=based_on,
                xml_element=style_elem
            )

            # Parse paragraph properties
            pPr = style_elem.find('w:pPr', namespaces=NAMESPACES)
            if pPr is not None:
                style.paragraph_formatting = self._parse_paragraph_formatting(pPr)

            # Parse run properties
            rPr = style_elem.find('w:rPr', namespaces=NAMESPACES)
            if rPr is not None:
                style.run_formatting = self._parse_run_formatting(rPr)

            styles[style_id] = style

        return styles

    def _parse_numbering(self) -> dict[str, NumberingDefinition]:
        """Parse numbering.xml and return numbering definitions."""
        definitions = {}
        abstract_nums = {}

        # First, parse abstract numbering definitions
        for abstract_elem in self.numbering_xml.xpath('//w:abstractNum', namespaces=NAMESPACES):
            abstract_id = abstract_elem.get(f'{{{NAMESPACES["w"]}}}abstractNumId')
            levels = {}

            for lvl_elem in abstract_elem.findall('w:lvl', namespaces=NAMESPACES):
                level_num = int(lvl_elem.get(f'{{{NAMESPACES["w"]}}}ilvl', 0))

                start_elem = lvl_elem.find('w:start', namespaces=NAMESPACES)
                start = int(start_elem.get(f'{{{NAMESPACES["w"]}}}val', 1)) if start_elem is not None else 1

                fmt_elem = lvl_elem.find('w:numFmt', namespaces=NAMESPACES)
                num_fmt = fmt_elem.get(f'{{{NAMESPACES["w"]}}}val', 'decimal') if fmt_elem is not None else 'decimal'

                text_elem = lvl_elem.find('w:lvlText', namespaces=NAMESPACES)
                level_text = text_elem.get(f'{{{NAMESPACES["w"]}}}val', '%1.') if text_elem is not None else '%1.'

                pPr = lvl_elem.find('w:pPr', namespaces=NAMESPACES)
                indent_left = None
                indent_hanging = None
                if pPr is not None:
                    ind = pPr.find('w:ind', namespaces=NAMESPACES)
                    if ind is not None:
                        indent_left = ind.get(f'{{{NAMESPACES["w"]}}}left')
                        indent_left = int(indent_left) if indent_left else None
                        indent_hanging = ind.get(f'{{{NAMESPACES["w"]}}}hanging')
                        indent_hanging = int(indent_hanging) if indent_hanging else None

                levels[level_num] = NumberingLevel(
                    level=level_num,
                    start=start,
                    num_fmt=num_fmt,
                    level_text=level_text,
                    indent_left=indent_left,
                    indent_hanging=indent_hanging
                )

            abstract_nums[abstract_id] = levels

        # Then, parse concrete numbering instances
        for num_elem in self.numbering_xml.xpath('//w:num', namespaces=NAMESPACES):
            num_id = num_elem.get(f'{{{NAMESPACES["w"]}}}numId')
            abstract_ref = num_elem.find('w:abstractNumId', namespaces=NAMESPACES)

            if abstract_ref is not None:
                abstract_id = abstract_ref.get(f'{{{NAMESPACES["w"]}}}val')
                definitions[num_id] = NumberingDefinition(
                    num_id=num_id,
                    abstract_num_id=abstract_id,
                    levels=abstract_nums.get(abstract_id, {})
                )

        return definitions

    def _parse_relationships(self, zf: zipfile.ZipFile) -> dict:
        """Parse document relationships."""
        rels = {}
        rels_xml = etree.parse(zf.open('word/_rels/document.xml.rels'))

        for rel in rels_xml.xpath('//*[local-name()="Relationship"]'):
            rel_id = rel.get('Id')
            rel_type = rel.get('Type')
            target = rel.get('Target')
            rels[rel_id] = {'type': rel_type, 'target': target}

        return rels

    def _parse_body(self, structure: DocumentStructure):
        """Parse the document body and populate structure."""
        body = self.document_xml.find('.//w:body', namespaces=NAMESPACES)
        if body is None:
            return

        para_index = 0
        table_index = 0

        for child in body:
            tag = etree.QName(child.tag).localname

            if tag == 'p':
                para = self._parse_paragraph(child)
                structure.paragraphs.append(para)
                structure.content_order.append(('paragraph', para_index))

                # Extract bookmarks
                self._extract_bookmarks(child, para_index, structure)

                # Extract cross-references
                self._extract_cross_references(child, structure)

                para_index += 1

            elif tag == 'tbl':
                table = self._parse_table(child)
                structure.tables.append(table)
                structure.content_order.append(('table', table_index))
                table_index += 1

    def _parse_paragraph(self, elem) -> Paragraph:
        """Parse a paragraph element."""
        para = Paragraph(xml_element=elem)

        # Parse paragraph properties
        pPr = elem.find('w:pPr', namespaces=NAMESPACES)
        if pPr is not None:
            para.formatting = self._parse_paragraph_formatting(pPr)

            # Extract numbering info
            numPr = pPr.find('w:numPr', namespaces=NAMESPACES)
            if numPr is not None:
                ilvl = numPr.find('w:ilvl', namespaces=NAMESPACES)
                numId = numPr.find('w:numId', namespaces=NAMESPACES)

                if ilvl is not None:
                    para.numbering_level = int(ilvl.get(f'{{{NAMESPACES["w"]}}}val', 0))
                if numId is not None:
                    para.numbering_id = numId.get(f'{{{NAMESPACES["w"]}}}val')

        # Parse runs
        for run_elem in elem.findall('w:r', namespaces=NAMESPACES):
            run = self._parse_run(run_elem)
            if run.text:  # Only add non-empty runs
                para.runs.append(run)

        return para

    def _parse_paragraph_formatting(self, pPr) -> ParagraphFormatting:
        """Parse paragraph formatting properties."""
        fmt = ParagraphFormatting()

        # Alignment
        jc = pPr.find('w:jc', namespaces=NAMESPACES)
        if jc is not None:
            fmt.alignment = jc.get(f'{{{NAMESPACES["w"]}}}val')

        # Indentation
        ind = pPr.find('w:ind', namespaces=NAMESPACES)
        if ind is not None:
            left = ind.get(f'{{{NAMESPACES["w"]}}}left')
            fmt.indent_left = int(left) if left else None

            right = ind.get(f'{{{NAMESPACES["w"]}}}right')
            fmt.indent_right = int(right) if right else None

            first_line = ind.get(f'{{{NAMESPACES["w"]}}}firstLine')
            fmt.indent_first_line = int(first_line) if first_line else None

            hanging = ind.get(f'{{{NAMESPACES["w"]}}}hanging')
            fmt.indent_hanging = int(hanging) if hanging else None

        # Spacing
        spacing = pPr.find('w:spacing', namespaces=NAMESPACES)
        if spacing is not None:
            before = spacing.get(f'{{{NAMESPACES["w"]}}}before')
            fmt.spacing_before = int(before) if before else None

            after = spacing.get(f'{{{NAMESPACES["w"]}}}after')
            fmt.spacing_after = int(after) if after else None

            line = spacing.get(f'{{{NAMESPACES["w"]}}}line')
            fmt.line_spacing = int(line) if line else None

            fmt.line_spacing_rule = spacing.get(f'{{{NAMESPACES["w"]}}}lineRule')

        # Outline level
        outline_lvl = pPr.find('w:outlineLvl', namespaces=NAMESPACES)
        if outline_lvl is not None:
            fmt.outline_level = int(outline_lvl.get(f'{{{NAMESPACES["w"]}}}val', 9))

        # Keep properties
        if pPr.find('w:keepNext', namespaces=NAMESPACES) is not None:
            fmt.keep_next = True
        if pPr.find('w:keepLines', namespaces=NAMESPACES) is not None:
            fmt.keep_lines = True
        if pPr.find('w:pageBreakBefore', namespaces=NAMESPACES) is not None:
            fmt.page_break_before = True

        # Style reference
        pStyle = pPr.find('w:pStyle', namespaces=NAMESPACES)
        if pStyle is not None:
            fmt.style_id = pStyle.get(f'{{{NAMESPACES["w"]}}}val')

        return fmt

    def _parse_run(self, elem) -> Run:
        """Parse a run element."""
        formatting = RunFormatting()

        # Parse run properties
        rPr = elem.find('w:rPr', namespaces=NAMESPACES)
        if rPr is not None:
            formatting = self._parse_run_formatting(rPr)

        # Get text content
        text_parts = []
        for child in elem:
            tag = etree.QName(child.tag).localname
            if tag == 't':
                text_parts.append(child.text or '')
            elif tag == 'tab':
                text_parts.append('\t')
            elif tag == 'br':
                text_parts.append('\n')

        return Run(
            text=''.join(text_parts),
            formatting=formatting,
            xml_element=elem
        )

    def _parse_run_formatting(self, rPr) -> RunFormatting:
        """Parse run formatting properties."""
        fmt = RunFormatting()

        # Bold
        b = rPr.find('w:b', namespaces=NAMESPACES)
        if b is not None:
            val = b.get(f'{{{NAMESPACES["w"]}}}val')
            fmt.bold = val != '0' and val != 'false'

        # Italic
        i = rPr.find('w:i', namespaces=NAMESPACES)
        if i is not None:
            val = i.get(f'{{{NAMESPACES["w"]}}}val')
            fmt.italic = val != '0' and val != 'false'

        # Underline
        u = rPr.find('w:u', namespaces=NAMESPACES)
        if u is not None:
            fmt.underline = u.get(f'{{{NAMESPACES["w"]}}}val')

        # Strike
        strike = rPr.find('w:strike', namespaces=NAMESPACES)
        if strike is not None:
            val = strike.get(f'{{{NAMESPACES["w"]}}}val')
            fmt.strike = val != '0' and val != 'false'

        # Font
        rFonts = rPr.find('w:rFonts', namespaces=NAMESPACES)
        if rFonts is not None:
            fmt.font_name = rFonts.get(f'{{{NAMESPACES["w"]}}}ascii') or rFonts.get(f'{{{NAMESPACES["w"]}}}hAnsi')

        # Font size
        sz = rPr.find('w:sz', namespaces=NAMESPACES)
        if sz is not None:
            fmt.font_size = int(sz.get(f'{{{NAMESPACES["w"]}}}val', 0))

        # Color
        color = rPr.find('w:color', namespaces=NAMESPACES)
        if color is not None:
            fmt.color = color.get(f'{{{NAMESPACES["w"]}}}val')

        # Highlight
        highlight = rPr.find('w:highlight', namespaces=NAMESPACES)
        if highlight is not None:
            fmt.highlight = highlight.get(f'{{{NAMESPACES["w"]}}}val')

        # Vertical alignment (superscript/subscript)
        vertAlign = rPr.find('w:vertAlign', namespaces=NAMESPACES)
        if vertAlign is not None:
            val = vertAlign.get(f'{{{NAMESPACES["w"]}}}val')
            fmt.superscript = val == 'superscript'
            fmt.subscript = val == 'subscript'

        # Caps
        if rPr.find('w:smallCaps', namespaces=NAMESPACES) is not None:
            fmt.small_caps = True
        if rPr.find('w:caps', namespaces=NAMESPACES) is not None:
            fmt.all_caps = True

        return fmt

    def _parse_table(self, elem) -> Table:
        """Parse a table element."""
        table = Table(xml_element=elem)

        # Parse table properties
        tblPr = elem.find('w:tblPr', namespaces=NAMESPACES)
        if tblPr is not None:
            # Width
            tblW = tblPr.find('w:tblW', namespaces=NAMESPACES)
            if tblW is not None:
                w = tblW.get(f'{{{NAMESPACES["w"]}}}w')
                table.width = int(w) if w else None

            # Alignment
            jc = tblPr.find('w:jc', namespaces=NAMESPACES)
            if jc is not None:
                table.alignment = jc.get(f'{{{NAMESPACES["w"]}}}val')

        # Parse rows
        for tr_elem in elem.findall('w:tr', namespaces=NAMESPACES):
            row = self._parse_table_row(tr_elem)
            table.rows.append(row)

        return table

    def _parse_table_row(self, elem) -> TableRow:
        """Parse a table row element."""
        row = TableRow(xml_element=elem)

        # Parse row properties
        trPr = elem.find('w:trPr', namespaces=NAMESPACES)
        if trPr is not None:
            trHeight = trPr.find('w:trHeight', namespaces=NAMESPACES)
            if trHeight is not None:
                h = trHeight.get(f'{{{NAMESPACES["w"]}}}val')
                row.height = int(h) if h else None

            if trPr.find('w:tblHeader', namespaces=NAMESPACES) is not None:
                row.is_header = True

        # Parse cells
        for tc_elem in elem.findall('w:tc', namespaces=NAMESPACES):
            cell = self._parse_table_cell(tc_elem)
            row.cells.append(cell)

        return row

    def _parse_table_cell(self, elem) -> TableCell:
        """Parse a table cell element."""
        cell = TableCell(xml_element=elem)

        # Parse cell properties
        tcPr = elem.find('w:tcPr', namespaces=NAMESPACES)
        if tcPr is not None:
            # Width
            tcW = tcPr.find('w:tcW', namespaces=NAMESPACES)
            if tcW is not None:
                w = tcW.get(f'{{{NAMESPACES["w"]}}}w')
                cell.width = int(w) if w else None

            # Column span
            gridSpan = tcPr.find('w:gridSpan', namespaces=NAMESPACES)
            if gridSpan is not None:
                cell.col_span = int(gridSpan.get(f'{{{NAMESPACES["w"]}}}val', 1))

            # Vertical alignment
            vAlign = tcPr.find('w:vAlign', namespaces=NAMESPACES)
            if vAlign is not None:
                cell.vertical_align = vAlign.get(f'{{{NAMESPACES["w"]}}}val')

        # Parse paragraphs in cell
        for p_elem in elem.findall('w:p', namespaces=NAMESPACES):
            para = self._parse_paragraph(p_elem)
            cell.paragraphs.append(para)

        return cell

    def _extract_bookmarks(self, elem, para_index: int, structure: DocumentStructure):
        """Extract bookmark definitions from a paragraph."""
        for bookmark_start in elem.findall('.//w:bookmarkStart', namespaces=NAMESPACES):
            name = bookmark_start.get(f'{{{NAMESPACES["w"]}}}name')
            bm_id = bookmark_start.get(f'{{{NAMESPACES["w"]}}}id')

            if name and not name.startswith('_'):  # Skip internal bookmarks
                structure.bookmarks[name] = Bookmark(
                    name=name,
                    bookmark_id=bm_id,
                    start_para_index=para_index,
                    end_para_index=para_index  # Will be updated when end is found
                )

    def _extract_cross_references(self, elem, structure: DocumentStructure):
        """Extract cross-references from a paragraph."""
        # Look for field codes (fldChar and instrText)
        field_codes = elem.findall('.//w:instrText', namespaces=NAMESPACES)

        for instr in field_codes:
            code = instr.text or ''
            code = code.strip()

            if code.startswith('REF ') or code.startswith('PAGEREF ') or code.startswith('SEQ '):
                ref_type = code.split()[0]
                bookmark_name = code.split()[1] if len(code.split()) > 1 else None

                # Clean up bookmark name (remove formatting flags)
                if bookmark_name:
                    bookmark_name = bookmark_name.split('\\')[0].strip()

                cross_ref = CrossReference(
                    field_code=code,
                    bookmark_name=bookmark_name,
                    ref_type=ref_type,
                    xml_element=elem
                )
                structure.cross_references.append(cross_ref)

    def get_raw_xml(self, part: str = 'document') -> etree._Element:
        """Get raw XML for a document part."""
        if part == 'document' and self.document_xml is not None:
            return copy.deepcopy(self.document_xml.getroot())
        elif part == 'styles' and self.styles_xml is not None:
            return copy.deepcopy(self.styles_xml.getroot())
        elif part == 'numbering' and self.numbering_xml is not None:
            return copy.deepcopy(self.numbering_xml.getroot())
        return None
