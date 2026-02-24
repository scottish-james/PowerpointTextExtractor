"""
Converts structured presentation data to clean markdown.
"""

import re


class MarkdownConverter:
    """
    Converts structured presentation data to clean markdown format using XML semantic roles.
    """

    def convert_structured_data_to_markdown(self, data, convert_slide_titles=True):
        """Convert an entire presentation to markdown. Returns a single markdown string."""
        markdown_parts = []

        for slide in data["slides"]:
            markdown_parts.append(f"\n<!-- Slide {slide['slide_number']} -->\n")

            for block in slide["content_blocks"]:
                block_markdown = None

                if block["type"] == "text":
                    block_markdown = self._convert_text_block_to_markdown(block)
                elif block["type"] == "table":
                    block_markdown = self._convert_table_to_markdown(block)
                elif block["type"] == "image":
                    block_markdown = self._convert_image_to_markdown(block)
                elif block["type"] == "chart":
                    block_markdown = self._convert_chart_to_markdown(block)
                elif block["type"] == "group":
                    block_markdown = self._convert_group_to_markdown(block)

                if block_markdown:
                    markdown_parts.append(block_markdown)

        return "\n\n".join(filter(None, markdown_parts))

    def _convert_text_block_to_markdown(self, block):
        """Convert a text block to markdown, using semantic role to set heading level."""
        lines = []
        semantic_role = block.get("semantic_role", "other")

        if semantic_role == "title":
            for para in block["paragraphs"]:
                if para.get("clean_text"):
                    formatted_text = self._build_formatted_text_from_runs(
                        para["formatted_runs"], para["clean_text"]
                    )
                    lines.append(f"# {formatted_text}")
        elif semantic_role == "subtitle":
            for para in block["paragraphs"]:
                if para.get("clean_text"):
                    formatted_text = self._build_formatted_text_from_runs(
                        para["formatted_runs"], para["clean_text"]
                    )
                    lines.append(f"## {formatted_text}")
        else:
            for para in block["paragraphs"]:
                line = self._convert_paragraph_to_markdown(para)
                if line:
                    lines.append(line)

        result = "\n".join(lines)

        if block.get("shape_hyperlink") and result:
            result = f"[{result}]({block['shape_hyperlink']})"

        return result

    def _convert_paragraph_to_markdown(self, para):
        """Convert a single paragraph to markdown with structure and formatting."""
        if not para.get("clean_text"):
            return ""

        formatted_text = self._build_formatted_text_from_runs(
            para["formatted_runs"], para["clean_text"]
        )

        hints = para.get("hints", {})

        if hints.get("is_bullet", False):
            level = max(hints.get("bullet_level", 0), 0)
            indent = "  " * level
            return f"{indent}- {formatted_text}"
        elif hints.get("is_numbered", False):
            return f"1. {formatted_text}"
        elif hints.get("likely_heading", False):
            if hints.get("all_caps") or len(para["clean_text"]) < 30:
                return f"## {formatted_text}"
            else:
                return f"### {formatted_text}"
        else:
            return formatted_text

    def _convert_group_to_markdown(self, block):
        """Convert a group of shapes to markdown by processing each child block."""
        extracted_blocks = block.get("extracted_blocks", [])

        if not extracted_blocks:
            return ""

        content_parts = []

        for extracted_block in extracted_blocks:
            content = None

            if extracted_block["type"] == "text":
                content = self._convert_text_block_to_markdown(extracted_block)
            elif extracted_block["type"] == "image":
                content = self._convert_image_to_markdown(extracted_block)
            elif extracted_block["type"] == "table":
                content = self._convert_table_to_markdown(extracted_block)
            elif extracted_block["type"] == "chart":
                content = self._convert_chart_to_markdown(extracted_block)
            elif extracted_block["type"] == "shape":
                content = f"[Shape: {extracted_block.get('shape_subtype', 'unknown')}]"

            if content:
                content_parts.append(content)

        group_md = "\n\n".join(content_parts) if content_parts else ""

        if block.get("hyperlink") and group_md:
            group_md = f"[{group_md}]({block['hyperlink']})"

        return group_md

    def _build_formatted_text_from_runs(self, runs, clean_text):
        """Build formatted markdown text from individual text runs."""
        if not runs:
            return clean_text

        text_runs = [run for run in runs if run.get("text")]

        if not text_runs:
            return clean_text

        all_bold = all(run.get("bold", False) for run in text_runs)
        all_italic = all(run.get("italic", False) for run in text_runs)
        all_have_hyperlinks = all(run.get("hyperlink") for run in text_runs)

        if all_have_hyperlinks:
            all_same_hyperlink = len(set(run.get("hyperlink") for run in text_runs)) == 1
        else:
            all_same_hyperlink = False

        if all_bold and all_italic and not all_same_hyperlink:
            return f"***{clean_text}***"
        elif all_bold and not all_same_hyperlink:
            return f"**{clean_text}**"
        elif all_italic and not all_same_hyperlink:
            return f"*{clean_text}*"
        elif all_same_hyperlink:
            hyperlink = text_runs[0]["hyperlink"]
            if all_bold and all_italic:
                return f"[***{clean_text}***]({hyperlink})"
            elif all_bold:
                return f"[**{clean_text}**]({hyperlink})"
            elif all_italic:
                return f"[*{clean_text}*]({hyperlink})"
            else:
                return f"[{clean_text}]({hyperlink})"

        formatted_parts = []
        for run in runs:
            text = run["text"]
            if not text:
                continue

            if run.get("bold") and run.get("italic"):
                text = f"***{text}***"
            elif run.get("bold"):
                text = f"**{text}**"
            elif run.get("italic"):
                text = f"*{text}*"

            if run.get("hyperlink"):
                text = f"[{text}]({run['hyperlink']})"

            formatted_parts.append(text)

        return "".join(formatted_parts)

    def _convert_table_to_markdown(self, block):
        """Convert table data to a markdown table."""
        if not block["data"]:
            return ""

        markdown = ""
        for i, row in enumerate(block["data"]):
            escaped_row = [cell.replace("|", "\\|") for cell in row]
            markdown += "| " + " | ".join(escaped_row) + " |\n"
            if i == 0:
                markdown += "| " + " | ".join("---" for _ in row) + " |\n"

        return markdown

    def _convert_image_to_markdown(self, block):
        """Convert an image block to markdown image syntax."""
        image_md = f"![{block['alt_text']}](image)"

        if block.get("hyperlink"):
            image_md = f"[{image_md}]({block['hyperlink']})"

        return image_md

    def _convert_chart_to_markdown(self, block):
        """Convert a chart block to markdown with a diagram candidate annotation."""
        chart_md = f"**Chart: {block.get('title', 'Untitled Chart')}**\n"
        chart_md += f"*Chart Type: {block.get('chart_type', 'unknown')}*\n\n"

        if block.get('categories') and block.get('series'):
            chart_md += "Data:\n"
            for series in block['series']:
                if series.get('name'):
                    chart_md += f"- {series['name']}: "
                    if series.get('values'):
                        chart_md += ", ".join(map(str, series['values'][:5]))
                        if len(series['values']) > 5:
                            chart_md += "..."
                    chart_md += "\n"

        if block.get("hyperlink"):
            chart_md = f"[{chart_md}]({block['hyperlink']})"

        return chart_md