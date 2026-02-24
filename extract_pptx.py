#!/usr/bin/env python3
"""
extract_pptx.py — Extract text from files as markdown.

Accepts a file as a BytesIO stream from a Flask request.
PowerPoint files use the full XML processing pipeline.
All other file types are handled by MarkItDown.

Usage:
    from extract_pptx import extract_to_markdown
    markdown = extract_to_markdown(file_bytes, filename)
"""

import sys
import os
from io import BytesIO

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(SCRIPT_DIR, "src")
sys.path.insert(0, SRC_DIR)

from powerpoint import convert_pptx_to_markdown_enhanced, process_powerpoint_file
from markitdown import MarkItDown

POWERPOINT_EXTENSIONS = {'.pptx', '.ppt'}

claud
def extract_to_markdown(file_bytes: BytesIO, filename: str) -> str:
    """
    Convert an uploaded file to markdown.

    Args:
        file_bytes: File contents as a BytesIO stream from a Flask request.
        filename: Original filename, used to determine file type.

    Returns:
        Markdown string.
    """
    ext = os.path.splitext(filename)[1].lower()

    if ext in POWERPOINT_EXTENSIONS:
        return convert_pptx_to_markdown_enhanced(file_bytes)
    else:
        return _convert_with_markitdown(file_bytes, filename)


def _convert_with_markitdown(file_bytes: BytesIO, filename: str) -> str:
    """Convert a non-PowerPoint file to markdown using MarkItDown."""
    md = MarkItDown()
    result = md.convert(file_bytes, filename=filename)

    try:
        return result.markdown
    except AttributeError:
        try:
            return result.text_content
        except AttributeError:
            raise Exception("MarkItDown returned an unrecognised result format.")