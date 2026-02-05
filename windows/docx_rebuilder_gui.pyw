#!/usr/bin/env python
"""
DOCX Rebuilder GUI Launcher for Windows.

This file has a .pyw extension so Windows runs it without a console window.
Double-click this file or associate it with .docx files.
"""

import sys
import os

# Add parent directory to path for imports
script_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(script_dir)
sys.path.insert(0, parent_dir)

from docx_rebuilder.gui import run_gui

if __name__ == '__main__':
    input_file = sys.argv[1] if len(sys.argv) > 1 else None
    run_gui(input_file)
