"""Tests for the numbering handler module."""

import pytest

from docx_rebuilder.numbering import NumberingHandler, NumberingState
from docx_rebuilder.parser import DocumentStructure, NumberingDefinition, NumberingLevel


def test_numbering_state_increment():
    """Test numbering state counter increment."""
    state = NumberingState(num_id="1")

    # First increment should return 1
    assert state.increment(0, start=1) == 1
    assert state.increment(0, start=1) == 2
    assert state.increment(0, start=1) == 3


def test_numbering_state_level_reset():
    """Test that lower levels are reset when higher level increments."""
    state = NumberingState(num_id="1")

    # Build up counts at multiple levels
    state.increment(0)  # 1
    state.increment(1)  # 1.1
    state.increment(1)  # 1.2

    # Increment level 0 should reset level 1
    state.increment(0)  # 2
    assert state.get_current(0) == 2
    assert state.get_current(1) == 0  # Reset


def test_format_number_decimal():
    """Test decimal number formatting."""
    structure = DocumentStructure()
    handler = NumberingHandler(structure)

    assert handler._format_number(1, 'decimal') == '1'
    assert handler._format_number(10, 'decimal') == '10'
    assert handler._format_number(100, 'decimal') == '100'


def test_format_number_letters():
    """Test letter number formatting."""
    structure = DocumentStructure()
    handler = NumberingHandler(structure)

    assert handler._format_number(1, 'lowerLetter') == 'a'
    assert handler._format_number(26, 'lowerLetter') == 'z'
    assert handler._format_number(27, 'lowerLetter') == 'aa'

    assert handler._format_number(1, 'upperLetter') == 'A'
    assert handler._format_number(26, 'upperLetter') == 'Z'


def test_format_number_roman():
    """Test Roman numeral formatting."""
    structure = DocumentStructure()
    handler = NumberingHandler(structure)

    assert handler._format_number(1, 'lowerRoman') == 'i'
    assert handler._format_number(4, 'lowerRoman') == 'iv'
    assert handler._format_number(9, 'lowerRoman') == 'ix'

    assert handler._format_number(1, 'upperRoman') == 'I'
    assert handler._format_number(10, 'upperRoman') == 'X'
    assert handler._format_number(50, 'upperRoman') == 'L'


def test_format_number_bullet():
    """Test bullet formatting."""
    structure = DocumentStructure()
    handler = NumberingHandler(structure)

    assert handler._format_number(1, 'bullet') == '•'
    assert handler._format_number(5, 'bullet') == '•'
