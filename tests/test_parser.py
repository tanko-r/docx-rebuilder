"""Tests for the DOCX parser module."""

import pytest
from pathlib import Path

from docx_rebuilder.parser import (
    NAMESPACES, RunFormatting, ParagraphFormatting,
    Run, Paragraph, Table
)


def test_run_formatting_defaults():
    """Test that RunFormatting has correct defaults."""
    fmt = RunFormatting()

    assert fmt.bold is False
    assert fmt.italic is False
    assert fmt.underline is None
    assert fmt.strike is False
    assert fmt.font_name is None
    assert fmt.font_size is None


def test_paragraph_formatting_defaults():
    """Test that ParagraphFormatting has correct defaults."""
    fmt = ParagraphFormatting()

    assert fmt.alignment is None
    assert fmt.indent_left is None
    assert fmt.indent_right is None
    assert fmt.style_id is None


def test_paragraph_text_property():
    """Test that Paragraph.text concatenates run text."""
    para = Paragraph()
    para.runs = [
        Run(text="Hello ", formatting=RunFormatting()),
        Run(text="World", formatting=RunFormatting()),
    ]

    assert para.text == "Hello World"


def test_namespaces_defined():
    """Test that required OOXML namespaces are defined."""
    assert 'w' in NAMESPACES
    assert 'r' in NAMESPACES
    assert NAMESPACES['w'] == 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
