"""
Formatting Analyzer - Detects style conventions and builds formatting models.

This module analyzes document formatting to detect patterns and infer
the intended style conventions, disregarding variations/inconsistencies.
"""

from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Optional
from statistics import mode, mean, stdev

from .parser import (
    DocumentStructure, Paragraph, ParagraphFormatting, RunFormatting,
    NumberingDefinition, Style
)


@dataclass
class FormattingCluster:
    """A cluster of similar formatting patterns."""
    formatting: ParagraphFormatting | RunFormatting
    frequency: int
    source_indices: list[int] = field(default_factory=list)


@dataclass
class StyleConvention:
    """Detected style convention for a particular purpose."""
    name: str
    paragraph_formatting: Optional[ParagraphFormatting] = None
    run_formatting: Optional[RunFormatting] = None
    numbering_level: Optional[int] = None
    usage_count: int = 0
    confidence: float = 0.0


@dataclass
class NumberingScheme:
    """Detected numbering scheme pattern."""
    scheme_id: str
    levels: dict[int, dict] = field(default_factory=dict)
    pattern_description: str = ""
    frequency: int = 0


@dataclass
class FormattingModel:
    """Complete model of detected formatting conventions."""
    heading_styles: dict[int, StyleConvention] = field(default_factory=dict)
    body_style: Optional[StyleConvention] = None
    numbering_schemes: list[NumberingScheme] = field(default_factory=list)
    indent_unit: int = 720  # Default: 0.5 inch in twips
    base_font_size: int = 24  # Default: 12pt in half-points
    base_font_name: str = "Times New Roman"
    default_alignment: str = "left"
    line_spacing: int = 240  # Default: single spacing
    spacing_after_paragraph: int = 0
    detected_styles: dict[str, StyleConvention] = field(default_factory=dict)


class FormattingAnalyzer:
    """Analyzes document formatting to detect style conventions."""

    def __init__(self, structure: DocumentStructure):
        self.structure = structure
        self.model = FormattingModel()

    def analyze(self) -> FormattingModel:
        """Perform full analysis and return formatting model."""
        self._analyze_fonts()
        self._analyze_paragraph_formatting()
        self._analyze_headings()
        self._analyze_numbering()
        self._analyze_indentation()
        self._analyze_existing_styles()
        return self.model

    def _analyze_fonts(self):
        """Analyze and detect base font settings."""
        font_names = Counter()
        font_sizes = Counter()

        for para in self.structure.paragraphs:
            for run in para.runs:
                if run.formatting.font_name:
                    font_names[run.formatting.font_name] += len(run.text)
                if run.formatting.font_size:
                    font_sizes[run.formatting.font_size] += len(run.text)

        # Find most common font (weighted by text length)
        if font_names:
            self.model.base_font_name = font_names.most_common(1)[0][0]

        if font_sizes:
            self.model.base_font_size = font_sizes.most_common(1)[0][0]

    def _analyze_paragraph_formatting(self):
        """Analyze paragraph formatting patterns."""
        alignments = Counter()
        line_spacings = Counter()
        spacing_afters = Counter()

        for para in self.structure.paragraphs:
            fmt = para.formatting

            if fmt.alignment:
                alignments[fmt.alignment] += 1

            if fmt.line_spacing:
                line_spacings[fmt.line_spacing] += 1

            if fmt.spacing_after is not None:
                spacing_afters[fmt.spacing_after] += 1

        # Determine most common values
        if alignments:
            self.model.default_alignment = alignments.most_common(1)[0][0]

        if line_spacings:
            self.model.line_spacing = line_spacings.most_common(1)[0][0]

        if spacing_afters:
            self.model.spacing_after_paragraph = spacing_afters.most_common(1)[0][0]

        # Build body style convention
        self.model.body_style = StyleConvention(
            name="Body",
            paragraph_formatting=ParagraphFormatting(
                alignment=self.model.default_alignment,
                line_spacing=self.model.line_spacing,
                spacing_after=self.model.spacing_after_paragraph
            ),
            run_formatting=RunFormatting(
                font_name=self.model.base_font_name,
                font_size=self.model.base_font_size
            ),
            confidence=0.9
        )

    def _analyze_headings(self):
        """Detect heading styles based on formatting patterns."""
        # Group paragraphs by potential heading indicators
        potential_headings = defaultdict(list)

        for i, para in enumerate(self.structure.paragraphs):
            # Check for heading indicators
            is_heading = False
            heading_level = None

            # Check outline level
            if para.formatting.outline_level is not None and para.formatting.outline_level < 9:
                is_heading = True
                heading_level = para.formatting.outline_level + 1

            # Check for style-based headings
            if para.formatting.style_id:
                style_id = para.formatting.style_id.lower()
                if 'heading' in style_id:
                    is_heading = True
                    # Try to extract level from style name
                    for j in range(1, 10):
                        if str(j) in style_id:
                            heading_level = j
                            break

            # Check for formatting-based headings (bold, larger font, short text)
            if not is_heading and para.text.strip():
                text_len = len(para.text.strip())
                has_bold = any(run.formatting.bold for run in para.runs)
                has_large_font = any(
                    run.formatting.font_size and run.formatting.font_size > self.model.base_font_size
                    for run in para.runs
                )

                # Short, bold text is likely a heading
                if text_len < 100 and has_bold and not para.text.strip().endswith('.'):
                    is_heading = True
                    if has_large_font:
                        heading_level = 1
                    else:
                        heading_level = 2

            if is_heading and heading_level:
                potential_headings[heading_level].append(i)

        # Build heading style conventions
        for level, indices in potential_headings.items():
            if indices:
                # Analyze common formatting for this heading level
                fmt = self._get_common_formatting(indices)
                run_fmt = self._get_common_run_formatting(indices)

                self.model.heading_styles[level] = StyleConvention(
                    name=f"Heading {level}",
                    paragraph_formatting=fmt,
                    run_formatting=run_fmt,
                    usage_count=len(indices),
                    confidence=min(1.0, len(indices) / 5)  # More occurrences = higher confidence
                )

    def _analyze_numbering(self):
        """Analyze numbering schemes and patterns."""
        numbering_usage = defaultdict(list)

        # Group paragraphs by numbering ID
        for i, para in enumerate(self.structure.paragraphs):
            if para.numbering_id:
                key = (para.numbering_id, para.numbering_level)
                numbering_usage[para.numbering_id].append((i, para.numbering_level))

        # Analyze each numbering scheme
        for num_id, usages in numbering_usage.items():
            if num_id in self.structure.numbering_definitions:
                num_def = self.structure.numbering_definitions[num_id]

                # Analyze the pattern
                levels_used = set(level for _, level in usages)
                level_info = {}

                for level in levels_used:
                    if level in num_def.levels:
                        level_def = num_def.levels[level]
                        level_info[level] = {
                            'format': level_def.num_fmt,
                            'text': level_def.level_text,
                            'start': level_def.start,
                            'indent': level_def.indent_left
                        }

                scheme = NumberingScheme(
                    scheme_id=num_id,
                    levels=level_info,
                    pattern_description=self._describe_numbering_pattern(level_info),
                    frequency=len(usages)
                )
                self.model.numbering_schemes.append(scheme)

    def _describe_numbering_pattern(self, levels: dict) -> str:
        """Generate a human-readable description of the numbering pattern."""
        descriptions = []

        format_names = {
            'decimal': 'numeric (1, 2, 3)',
            'lowerLetter': 'lowercase letters (a, b, c)',
            'upperLetter': 'uppercase letters (A, B, C)',
            'lowerRoman': 'lowercase roman (i, ii, iii)',
            'upperRoman': 'uppercase roman (I, II, III)',
            'bullet': 'bullet points',
        }

        for level in sorted(levels.keys()):
            info = levels[level]
            fmt_name = format_names.get(info['format'], info['format'])
            descriptions.append(f"Level {level + 1}: {fmt_name}")

        return '; '.join(descriptions)

    def _analyze_indentation(self):
        """Detect the standard indentation unit used in the document."""
        indent_values = []

        for para in self.structure.paragraphs:
            if para.formatting.indent_left:
                indent_values.append(para.formatting.indent_left)
            if para.formatting.indent_first_line:
                indent_values.append(para.formatting.indent_first_line)

        # Also check numbering definitions
        for num_def in self.structure.numbering_definitions.values():
            for level_def in num_def.levels.values():
                if level_def.indent_left:
                    indent_values.append(level_def.indent_left)

        if indent_values:
            # Find the GCD-like common unit
            # Common indent units: 720 (0.5"), 360 (0.25"), 1440 (1")
            self.model.indent_unit = self._find_indent_unit(indent_values)

    def _find_indent_unit(self, values: list[int]) -> int:
        """Find the likely base indentation unit from a list of indent values."""
        if not values:
            return 720  # Default 0.5"

        # Common Word indent units in twips
        common_units = [180, 360, 720, 1440]  # 0.125", 0.25", 0.5", 1"

        # Count how many values are divisible by each unit
        best_unit = 720
        best_count = 0

        for unit in common_units:
            count = sum(1 for v in values if v % unit == 0 or abs(v % unit) < 20)
            if count > best_count:
                best_count = count
                best_unit = unit

        return best_unit

    def _analyze_existing_styles(self):
        """Analyze existing document styles."""
        for style_id, style in self.structure.styles.items():
            if style.style_type == 'paragraph':
                convention = StyleConvention(
                    name=style.name,
                    paragraph_formatting=style.paragraph_formatting,
                    run_formatting=style.run_formatting
                )
                self.model.detected_styles[style_id] = convention

    def _get_common_formatting(self, indices: list[int]) -> ParagraphFormatting:
        """Get the most common paragraph formatting from a list of paragraph indices."""
        alignments = Counter()
        indent_lefts = []
        spacing_befores = []
        spacing_afters = []

        for i in indices:
            if i < len(self.structure.paragraphs):
                fmt = self.structure.paragraphs[i].formatting

                if fmt.alignment:
                    alignments[fmt.alignment] += 1
                if fmt.indent_left is not None:
                    indent_lefts.append(fmt.indent_left)
                if fmt.spacing_before is not None:
                    spacing_befores.append(fmt.spacing_before)
                if fmt.spacing_after is not None:
                    spacing_afters.append(fmt.spacing_after)

        result = ParagraphFormatting()

        if alignments:
            result.alignment = alignments.most_common(1)[0][0]

        if indent_lefts:
            result.indent_left = self._get_most_common_value(indent_lefts)

        if spacing_befores:
            result.spacing_before = self._get_most_common_value(spacing_befores)

        if spacing_afters:
            result.spacing_after = self._get_most_common_value(spacing_afters)

        return result

    def _get_common_run_formatting(self, indices: list[int]) -> RunFormatting:
        """Get the most common run formatting from a list of paragraph indices."""
        bold_count = 0
        italic_count = 0
        font_sizes = Counter()
        font_names = Counter()
        total_runs = 0

        for i in indices:
            if i < len(self.structure.paragraphs):
                for run in self.structure.paragraphs[i].runs:
                    total_runs += 1
                    if run.formatting.bold:
                        bold_count += 1
                    if run.formatting.italic:
                        italic_count += 1
                    if run.formatting.font_size:
                        font_sizes[run.formatting.font_size] += 1
                    if run.formatting.font_name:
                        font_names[run.formatting.font_name] += 1

        result = RunFormatting()

        if total_runs > 0:
            result.bold = bold_count > total_runs / 2
            result.italic = italic_count > total_runs / 2

        if font_sizes:
            result.font_size = font_sizes.most_common(1)[0][0]

        if font_names:
            result.font_name = font_names.most_common(1)[0][0]

        return result

    def _get_most_common_value(self, values: list[int], tolerance: int = 20) -> int:
        """Get the most common value, grouping similar values together."""
        if not values:
            return 0

        # Group similar values
        groups = defaultdict(list)
        for v in values:
            # Find existing group or create new one
            found = False
            for key in groups:
                if abs(v - key) <= tolerance:
                    groups[key].append(v)
                    found = True
                    break
            if not found:
                groups[v].append(v)

        # Find largest group
        largest_group = max(groups.values(), key=len)
        return int(mean(largest_group))

    def get_recommended_style_for_paragraph(self, para: Paragraph) -> str:
        """Recommend the best style to apply to a paragraph."""
        # Check if it matches a heading pattern
        for level, convention in self.model.heading_styles.items():
            if self._matches_convention(para, convention):
                return f"Heading{level}"

        # Check for numbered paragraph
        if para.numbering_id:
            return "ListParagraph"

        # Default to body style
        return "Normal"

    def _matches_convention(self, para: Paragraph, convention: StyleConvention) -> bool:
        """Check if a paragraph matches a style convention."""
        # Simple matching based on formatting similarity
        if convention.paragraph_formatting:
            if (para.formatting.alignment and convention.paragraph_formatting.alignment and
                para.formatting.alignment != convention.paragraph_formatting.alignment):
                return False

        if convention.run_formatting and para.runs:
            # Check if majority of runs match
            matching_runs = 0
            for run in para.runs:
                if (convention.run_formatting.bold == run.formatting.bold and
                    convention.run_formatting.italic == run.formatting.italic):
                    matching_runs += 1

            if matching_runs < len(para.runs) / 2:
                return False

        return True

    def suggest_corrections(self) -> list[dict]:
        """Suggest formatting corrections based on detected conventions."""
        suggestions = []

        for i, para in enumerate(self.structure.paragraphs):
            # Check for inconsistent indentation
            if para.formatting.indent_left:
                expected = self._round_to_unit(para.formatting.indent_left, self.model.indent_unit)
                if para.formatting.indent_left != expected:
                    suggestions.append({
                        'type': 'indent_correction',
                        'paragraph_index': i,
                        'current': para.formatting.indent_left,
                        'suggested': expected,
                        'message': f'Indent {para.formatting.indent_left} should be {expected}'
                    })

            # Check for font inconsistencies
            for j, run in enumerate(para.runs):
                if run.formatting.font_name and run.formatting.font_name != self.model.base_font_name:
                    # Only suggest if this isn't a heading
                    recommended_style = self.get_recommended_style_for_paragraph(para)
                    if not recommended_style.startswith('Heading'):
                        suggestions.append({
                            'type': 'font_correction',
                            'paragraph_index': i,
                            'run_index': j,
                            'current': run.formatting.font_name,
                            'suggested': self.model.base_font_name,
                            'message': f'Font {run.formatting.font_name} should be {self.model.base_font_name}'
                        })

        return suggestions

    def _round_to_unit(self, value: int, unit: int) -> int:
        """Round a value to the nearest multiple of the unit."""
        return round(value / unit) * unit
