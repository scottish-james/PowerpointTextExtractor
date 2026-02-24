#!/usr/bin/env python3
"""
extract_pptx.py — Extract text from PowerPoint files as markdown.

Usage:
    python extract_pptx.py presentation.pptx
    python extract_pptx.py presentation.pptx -o output.md
    python extract_pptx.py presentation.pptx --summary
"""

import argparse
import sys
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(SCRIPT_DIR, "src")
sys.path.insert(0, SRC_DIR)

from powerpoint import convert_pptx_to_markdown_enhanced, process_powerpoint_file


def main():
    parser = argparse.ArgumentParser(
        description="Extract text from a PowerPoint file as markdown."
    )
    parser.add_argument("file", help="Path to the .pptx file")
    parser.add_argument("-o", "--output", help="Write output to this file instead of the default")
    parser.add_argument("--summary", action="store_true", help="Print processing summary with metadata")
    args = parser.parse_args()

    if not os.path.isfile(args.file):
        print(f"Error: '{args.file}' not found.", file=sys.stderr)
        sys.exit(1)

    # Default output: same name and location as input, with .md extension
    output_path = args.output or os.path.splitext(args.file)[0] + ".md"

    if args.summary:
        result = process_powerpoint_file(args.file, output_format="markdown")
        markdown = result["content"]
        method = result.get("processing_method", "unknown")
        metadata = result.get("metadata", {})

        print(f"Processing method: {method}")
        if metadata:
            print(f"Slide count: {metadata.get('slide_count', 'N/A')}")
            print(f"Title: {metadata.get('title', 'N/A')}")
        print("---")
    else:
        markdown = convert_pptx_to_markdown_enhanced(args.file)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(markdown)
    print(f"Written to {output_path}")


if __name__ == "__main__":
    main()