"""
Manifest Builder - Generates rebuild summary reports.

Shared between the CLI and GUI to produce consistent manifest output.
"""

from datetime import datetime
from pathlib import Path
from typing import Optional


def build_manifest(input_path: Path,
                   output_path: Path,
                   report: dict,
                   suggestions: list,
                   options: Optional[dict] = None) -> str:
    """Build a manifest text report describing the rebuild work performed.

    Args:
        input_path: Path to the original document
        output_path: Path to the rebuilt document
        report: Analysis report dict from DocxRebuilder.get_analysis_report()
        suggestions: List of correction suggestions from FormattingAnalyzer
        options: Optional dict of rebuild options that were applied

    Returns:
        Formatted manifest string
    """
    lines = []
    lines.append("=" * 55)
    lines.append("  DOCX REBUILDER - CONVERSION MANIFEST")
    lines.append(f"  Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 55)
    lines.append("")

    # File Info
    lines.append("INPUT / OUTPUT")
    lines.append("-" * 40)
    lines.append(f"  Input:   {input_path.name}")
    lines.append(f"  Output:  {output_path.name}")
    lines.append(f"  Folder:  {output_path.parent}")
    lines.append("")

    # Document Statistics
    lines.append("DOCUMENT STATISTICS")
    lines.append("-" * 40)
    para_count = report.get('paragraph_count', 0)
    table_count = report.get('table_count', 0)
    bookmark_count = report.get('bookmark_count', 0)
    xref_count = report.get('cross_reference_count', 0)

    lines.append(f"  Paragraphs processed:    {para_count}")
    lines.append(f"  Tables preserved:        {table_count}")
    lines.append(f"  Bookmarks found:         {bookmark_count}")
    lines.append(f"  Cross-references:        {xref_count}")
    lines.append("")

    # Headers/Footers
    headers = report.get('headers_count', 0)
    footers = report.get('footers_count', 0)
    sections = report.get('sections_count', 0)
    page_fields = report.get('page_number_fields', 0)

    lines.append("HEADERS, FOOTERS & PAGE NUMBERING")
    lines.append("-" * 40)
    lines.append(f"  Headers preserved:       {headers}")
    lines.append(f"  Footers preserved:       {footers}")
    lines.append(f"  Document sections:       {sections}")
    lines.append(f"  Page number fields:      {page_fields}")
    lines.append("")

    # Formatting Analysis
    lines.append("FORMATTING ANALYSIS")
    lines.append("-" * 40)
    base_font = report.get('base_font', 'Unknown')
    base_size = report.get('base_font_size', 0)
    alignment = report.get('default_alignment', 'left')
    indent = report.get('indent_unit', 720)
    heading_count = report.get('heading_styles_detected', 0)
    num_schemes = report.get('numbering_schemes', 0)

    # Translate alignment for lawyers
    align_names = {'left': 'Left', 'center': 'Center', 'right': 'Right',
                   'both': 'Justified', 'justify': 'Justified'}
    alignment_display = align_names.get(alignment, alignment)

    # Translate indent twips to inches
    indent_inches = indent / 1440

    lines.append(f"  Base font detected:      {base_font}")
    lines.append(f"  Base font size:          {base_size}pt")
    lines.append(f"  Default alignment:       {alignment_display}")
    lines.append(f"  Indent unit:             {indent_inches:.2f}\" ({indent} twips)")
    lines.append(f"  Heading styles found:    {heading_count}")
    lines.append(f"  Numbering schemes:       {num_schemes}")
    lines.append("")

    # Rebuild options applied
    if options is None:
        options = {
            'normalize_numbering': True,
            'normalize_styles': True,
            'preserve_cross_refs': True,
        }

    lines.append("ACTIONS PERFORMED")
    lines.append("-" * 40)
    lines.append("  [x] Parsed document structure")
    lines.append("  [x] Analyzed formatting conventions")

    if options.get('normalize_numbering', True):
        lines.append("  [x] Normalized numbering schemes")
    else:
        lines.append("  [ ] Numbering normalization (skipped)")

    if options.get('normalize_styles', True):
        lines.append("  [x] Normalized style consistency")
    else:
        lines.append("  [ ] Style normalization (skipped)")

    if options.get('preserve_cross_refs', True):
        lines.append("  [x] Preserved cross-references & bookmarks")
    else:
        lines.append("  [ ] Cross-reference preservation (skipped)")

    if headers or footers:
        lines.append("  [x] Preserved headers and footers")
    if page_fields:
        lines.append("  [x] Preserved page numbering")
    if table_count:
        lines.append("  [x] Preserved table formatting")

    lines.append("  [x] Rebuilt clean OOXML structure")
    lines.append("  [x] Maintained character formatting (bold, italic, underline)")
    lines.append("  [x] Preserved paragraph alignment and indentation")
    lines.append("")

    # Corrections applied
    if suggestions:
        indent_fixes = [s for s in suggestions if s.get('type') == 'indent_correction']
        font_fixes = [s for s in suggestions if s.get('type') == 'font_correction']
        other_fixes = [s for s in suggestions
                       if s.get('type') not in ('indent_correction', 'font_correction')]

        lines.append("FORMATTING CORRECTIONS APPLIED")
        lines.append("-" * 40)

        if indent_fixes:
            lines.append(f"  Indentation normalized:  {len(indent_fixes)} instances")
        if font_fixes:
            lines.append(f"  Font inconsistencies:    {len(font_fixes)} instances")
        if other_fixes:
            lines.append(f"  Other corrections:       {len(other_fixes)} instances")

        lines.append("")
        lines.append("  Details:")
        for s in suggestions[:10]:
            lines.append(f"    - {s.get('message', 'Unknown correction')}")
        if len(suggestions) > 10:
            lines.append(f"    ... and {len(suggestions) - 10} more corrections")
        lines.append("")
    else:
        lines.append("FORMATTING CORRECTIONS")
        lines.append("-" * 40)
        lines.append("  No inconsistencies detected.")
        lines.append("")

    # Footer
    lines.append("=" * 55)
    lines.append("  The rebuilt document should now have consistent")
    lines.append("  formatting that works cleanly with MS Word.")
    lines.append("=" * 55)

    return "\n".join(lines)
