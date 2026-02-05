"""
DOCX Rebuilder - Magically fixes Word document formatting for lawyers.

This tool analyzes document formatting, detects style conventions,
and rebuilds the OOXML with clean, consistent formatting.
"""

__version__ = "1.0.0"

from .rebuilder import DocxRebuilder
from .analyzer import FormattingAnalyzer
from .parser import DocxParser

__all__ = ["DocxRebuilder", "FormattingAnalyzer", "DocxParser"]
