"""
Handles text extraction from PowerPoint shapes as part of the processing pipeline.

Sits between ContentExtractor and MarkdownConverter. ContentExtractor routes
text-bearing shapes here, and the output is passed on to MarkdownConverter
for rendering.

Uses PowerPoint's internal XML to detect bullets, numbering and formatting
rather than guessing from text patterns. Each paragraph is returned as a
structured dict containing the cleaned text, per-run formatting (bold, italic,
hyperlinks) and hints that MarkdownConverter uses to apply the right markdown.
"""

import re


class TextProcessor:
    """
    Extracts structured text data from PowerPoint shapes.
    """

    def extract_text_frame(self, text_frame, shape):
        """Extract all paragraphs from a TextFrame, skipping empty ones."""
        if not text_frame.paragraphs:
            return None

        block = {
            "type": "text",
            "paragraphs": [],
            "shape_hyperlink": self._extract_shape_hyperlink(shape)
        }

        for para in text_frame.paragraphs:
            if not para.text.strip():
                continue
            para_data = self.process_paragraph(para)
            if para_data:
                block["paragraphs"].append(para_data)

        return block if block["paragraphs"] else None

    def extract_plain_text(self, shape):
        """Fallback for shapes that expose only a .text attribute with no paragraph structure."""
        if not hasattr(shape, 'text') or not shape.text:
            return None

        return {
            "type": "text",
            "paragraphs": [{
                "raw_text": shape.text,
                "clean_text": shape.text.strip(),
                "formatted_runs": [{"text": shape.text, "bold": False, "italic": False, "hyperlink": None}],
                "hints": self._analyze_plain_text_hints(shape.text)
            }],
            "shape_hyperlink": self._extract_shape_hyperlink(shape)
        }

    def process_paragraph(self, para):
        """
        Process a single paragraph using XML to detect bullets, numbering and level.
        Returns a dict with raw text, clean text, formatted runs and hints.
        """
        raw_text = para.text
        if not raw_text.strip():
            return None

        ppt_level = getattr(para, 'level', None)
        is_ppt_bullet, xml_level = self._check_xml_bullet_formatting(para)
        bullet_level = self._determine_bullet_level(is_ppt_bullet, xml_level, ppt_level)

        clean_text = raw_text.strip()
        if bullet_level >= 0:
            clean_text = self._remove_bullet_char(clean_text)

        formatted_runs = self._extract_runs_with_formatting(para.runs, clean_text, bullet_level >= 0)

        return {
            "raw_text": raw_text,
            "clean_text": clean_text,
            "formatted_runs": formatted_runs,
            "hints": {
                "has_powerpoint_level": ppt_level is not None,
                "powerpoint_level": ppt_level,
                "bullet_level": bullet_level,
                "is_bullet": bullet_level >= 0,
                "is_numbered": self._is_numbered_from_xml(para),
                "short_text": len(clean_text) < 100,
                "all_caps": clean_text.isupper() if clean_text else False,
            }
        }

    def _check_xml_bullet_formatting(self, para):
        """Check paragraph XML for bullet indicators and return (is_bullet, level)."""
        is_ppt_bullet = False
        xml_level = None

        try:
            if hasattr(para, '_p') and para._p is not None:
                xml_str = str(para._p.xml)

                if any(indicator in xml_str for indicator in ['buChar', 'buAutoNum', 'buFont']):
                    is_ppt_bullet = True

                    level_match = re.search(r'lvl="(\d+)"', xml_str)
                    if level_match:
                        xml_level = int(level_match.group(1))
        except:
            pass

        return is_ppt_bullet, xml_level

    def _is_numbered_from_xml(self, para):
        """Return True if the paragraph XML indicates an auto-numbered list."""
        try:
            if hasattr(para, '_p') and para._p is not None:
                return 'buAutoNum' in str(para._p.xml)
        except:
            pass
        return False

    def _determine_bullet_level(self, is_ppt_bullet, xml_level, ppt_level):
        """
        Resolve bullet level from available sources.
        Prefers XML level, falls back to paragraph level, returns -1 for non-bullets.
        """
        if is_ppt_bullet:
            return xml_level if xml_level is not None else (ppt_level if ppt_level is not None else 0)
        elif ppt_level is not None:
            return ppt_level
        return -1

    def _extract_runs_with_formatting(self, runs, clean_text, has_prefix_removed):
        """Extract bold, italic and hyperlink state from each text run."""
        if not runs:
            return [{"text": clean_text, "bold": False, "italic": False, "hyperlink": None}]

        formatted_runs = []

        if has_prefix_removed:
            full_text = "".join(run.text for run in runs)
            start_pos = self._find_clean_text_start_position(full_text, clean_text)

            char_count = 0
            for run in runs:
                run_text = run.text
                run_start = char_count
                run_end = char_count + len(run_text)

                if run_end <= start_pos:
                    char_count += len(run_text)
                    continue

                if run_start < start_pos < run_end:
                    run_text = run_text[start_pos - run_start:]

                if run_text:
                    formatted_runs.append(self._extract_run_formatting(run, run_text))

                char_count += len(run.text)
        else:
            for run in runs:
                if run.text:
                    formatted_runs.append(self._extract_run_formatting(run, run.text))

        return formatted_runs

    def _find_clean_text_start_position(self, full_text, clean_text):
        """Find where clean text begins within the original text after bullet removal."""
        for i in range(len(full_text)):
            if full_text[i:].strip() == clean_text:
                return i
        return 0

    def _extract_run_formatting(self, run, text_override=None):
        """Return bold, italic and hyperlink state for a single text run."""
        run_data = {
            "text": text_override if text_override is not None else run.text,
            "bold": False,
            "italic": False,
            "hyperlink": None
        }

        try:
            font = run.font
            if hasattr(font, 'bold') and font.bold:
                run_data["bold"] = True
            if hasattr(font, 'italic') and font.italic:
                run_data["italic"] = True
        except:
            pass

        try:
            if hasattr(run, 'hyperlink') and run.hyperlink and run.hyperlink.address:
                run_data["hyperlink"] = self._fix_url(run.hyperlink.address)
        except:
            pass

        return run_data

    def _remove_bullet_char(self, text):
        """Strip leading bullet characters from text."""
        if not text:
            return text
        return re.sub(r'^[•◦▪▫‣·○■□→►✓✗\-\*\+※◆◇]\s*', '', text)

    def _analyze_plain_text_hints(self, text):
        """Return basic hints for shapes without full paragraph structure."""
        if not text:
            return {}

        stripped = text.strip()

        return {
            "has_powerpoint_level": False,
            "powerpoint_level": None,
            "bullet_level": -1,
            "is_bullet": False,
            "is_numbered": False,
            "starts_with_bullet": False,
            "starts_with_number": False,
            "short_text": len(stripped) < 100,
            "all_caps": stripped.isupper() if stripped else False,
            "likely_heading": 0 < len(stripped) < 80
        }

    def _extract_shape_hyperlink(self, shape):
        """Return the URL if the entire shape has a clickable hyperlink, otherwise None."""
        try:
            if hasattr(shape, 'click_action') and shape.click_action:
                if hasattr(shape.click_action, 'hyperlink') and shape.click_action.hyperlink:
                    if shape.click_action.hyperlink.address:
                        return self._fix_url(shape.click_action.hyperlink.address)
        except:
            pass
        return None

    def _fix_url(self, url):
        """Add missing scheme to URLs and mailto: prefix to email addresses."""
        if not url:
            return url

        if '@' in url and not url.startswith('mailto:'):
            return f"mailto:{url}"

        if not url.startswith(('http://', 'https://', 'mailto:', 'tel:', 'ftp://', '#')):
            if url.startswith('www.') or any(
                    domain in url.lower() for domain in ['.com', '.org', '.net', '.edu', '.gov', '.io']):
                return f"https://{url}"

        return url