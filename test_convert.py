"""
Quick test script. Run from the project root:
    python test_convert.py path/to/file.pptx
"""

import sys
import os
from io import BytesIO

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(SCRIPT_DIR, "src")
sys.path.insert(0, SRC_DIR)

from extract_pptx import extract_to_markdown

file_path = sys.argv[1]
filename = os.path.basename(file_path)

with open(file_path, "rb") as f:
    file_bytes = BytesIO(f.read())

markdown = extract_to_markdown(file_bytes, filename)

output_path = os.path.splitext(file_path)[0] + ".md"
with open(output_path, "w", encoding="utf-8") as f:
    f.write(markdown)

print(f"Written to {output_path}")