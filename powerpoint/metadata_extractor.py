"""
Extracts metadata from PowerPoint files and embeds it into markdown output as an HTML comment.
"""

import os
from datetime import datetime


class MetadataExtractor:
    """
    Extracts and formats PowerPoint metadata.
    """

    def extract_pptx_metadata(self, presentation, file_path):
        """Extract all available metadata from a PowerPoint file."""
        metadata = {}

        try:
            core_props = presentation.core_properties

            metadata['filename'] = os.path.basename(file_path)
            metadata['file_size'] = os.path.getsize(file_path) if os.path.exists(file_path) else None

            metadata.update(self._extract_document_properties(core_props))
            metadata.update(self._extract_date_properties(core_props))
            metadata.update(self._extract_revision_properties(core_props))
            metadata.update(self._extract_presentation_properties(presentation))
            metadata.update(self._extract_application_properties(presentation))

        except Exception:
            pass

        return metadata

    def _extract_document_properties(self, core_props):
        """Extract title, author, subject, keywords and related fields."""
        return {
            'title': getattr(core_props, 'title', '') or '',
            'author': getattr(core_props, 'author', '') or '',
            'subject': getattr(core_props, 'subject', '') or '',
            'keywords': getattr(core_props, 'keywords', '') or '',
            'comments': getattr(core_props, 'comments', '') or '',
            'category': getattr(core_props, 'category', '') or '',
            'content_status': getattr(core_props, 'content_status', '') or '',
            'language': getattr(core_props, 'language', '') or '',
            'version': getattr(core_props, 'version', '') or '',
        }

    def _extract_date_properties(self, core_props):
        """Extract created, modified and last printed dates."""
        return {
            'created': getattr(core_props, 'created', None),
            'modified': getattr(core_props, 'modified', None),
            'last_modified_by': getattr(core_props, 'last_modified_by', '') or '',
            'last_printed': getattr(core_props, 'last_printed', None),
        }

    def _extract_revision_properties(self, core_props):
        """Extract revision number and identifier."""
        return {
            'revision': getattr(core_props, 'revision', None),
            'identifier': getattr(core_props, 'identifier', '') or '',
        }

    def _extract_presentation_properties(self, presentation):
        """Extract slide count, master count and layout types."""
        metadata = {
            'slide_count': len(presentation.slides)
        }

        try:
            slide_masters = presentation.slide_masters
            if slide_masters:
                metadata['slide_master_count'] = len(slide_masters)

                layout_names = []
                for master in slide_masters:
                    for layout in master.slide_layouts:
                        if hasattr(layout, 'name') and layout.name:
                            layout_names.append(layout.name)

                metadata['layout_types'] = ', '.join(set(layout_names)) if layout_names else ''
            else:
                metadata['slide_master_count'] = 0
                metadata['layout_types'] = ''
        except Exception:
            metadata['slide_master_count'] = 0
            metadata['layout_types'] = ''

        return metadata

    def _extract_application_properties(self, presentation):
        """Extract creating application, version and company."""
        metadata = {
            'application': '',
            'app_version': '',
            'company': '',
            'doc_security': None
        }

        try:
            app_props = presentation.app_properties if hasattr(presentation, 'app_properties') else None
            if app_props:
                metadata['application'] = getattr(app_props, 'application', '') or ''
                metadata['app_version'] = getattr(app_props, 'app_version', '') or ''
                metadata['company'] = getattr(app_props, 'company', '') or ''
                metadata['doc_security'] = getattr(app_props, 'doc_security', None)
        except Exception:
            pass

        return metadata

    def add_pptx_metadata(self, markdown_content, metadata):
        """Prepend metadata to markdown content as an HTML comment block."""
        metadata_comments = "\n<!-- POWERPOINT METADATA:\n"

        metadata_comments += self._format_document_metadata(metadata)
        metadata_comments += self._format_date_metadata(metadata)
        metadata_comments += self._format_file_metadata(metadata)
        metadata_comments += self._format_presentation_metadata(metadata)

        metadata_comments += "-->\n"

        return metadata_comments + markdown_content

    def _format_document_metadata(self, metadata):
        """Format document fields, skipping any that are empty."""
        formatted = ""

        if metadata.get('title'):
            formatted += f"Document Title: {metadata['title']}\n"
        if metadata.get('author'):
            formatted += f"Author: {metadata['author']}\n"
        if metadata.get('subject'):
            formatted += f"Subject: {metadata['subject']}\n"
        if metadata.get('keywords'):
            formatted += f"Keywords: {metadata['keywords']}\n"
        if metadata.get('category'):
            formatted += f"Category: {metadata['category']}\n"
        if metadata.get('comments'):
            formatted += f"Document Comments: {metadata['comments']}\n"
        if metadata.get('content_status'):
            formatted += f"Content Status: {metadata['content_status']}\n"
        if metadata.get('language'):
            formatted += f"Language: {metadata['language']}\n"
        if metadata.get('version'):
            formatted += f"Version: {metadata['version']}\n"

        return formatted

    def _format_date_metadata(self, metadata):
        """Format date fields, skipping any that are not set."""
        formatted = ""

        if metadata.get('created'):
            formatted += f"Created Date: {metadata['created']}\n"
        if metadata.get('modified'):
            formatted += f"Last Modified: {metadata['modified']}\n"
        if metadata.get('last_modified_by'):
            formatted += f"Last Modified By: {metadata['last_modified_by']}\n"
        if metadata.get('last_printed'):
            formatted += f"Last Printed: {metadata['last_printed']}\n"

        return formatted

    def _format_file_metadata(self, metadata):
        """Format file name, size and application fields."""
        formatted = ""

        formatted += f"Filename: {metadata.get('filename', 'unknown')}\n"

        if metadata.get('file_size'):
            file_size_mb = metadata['file_size'] / (1024 * 1024)
            formatted += f"File Size: {file_size_mb:.2f} MB\n"

        if metadata.get('application'):
            formatted += f"Created With: {metadata['application']}\n"
        if metadata.get('company'):
            formatted += f"Company: {metadata['company']}\n"

        return formatted

    def _format_presentation_metadata(self, metadata):
        """Format slide count, master count and layout types."""
        formatted = ""

        formatted += f"Slide Count: {metadata.get('slide_count', 0)}\n"

        if metadata.get('slide_master_count'):
            formatted += f"Slide Masters: {metadata['slide_master_count']}\n"
        if metadata.get('layout_types'):
            formatted += f"Layout Types: {metadata['layout_types']}\n"

        return formatted

    def get_metadata_summary(self, metadata):
        """Return a summary dict with key indicators for quick assessment."""
        summary = {
            'has_title': bool(metadata.get('title')),
            'has_author': bool(metadata.get('author')),
            'slide_count': metadata.get('slide_count', 0),
            'file_size_mb': None,
            'creation_date': metadata.get('created'),
            'last_modified': metadata.get('modified'),
            'has_keywords': bool(metadata.get('keywords')),
            'application': metadata.get('application', 'Unknown'),
        }

        if metadata.get('file_size'):
            summary['file_size_mb'] = round(metadata['file_size'] / (1024 * 1024), 2)

        return summary

    def validate_metadata(self, metadata):
        """
        Score metadata completeness and return any issues or recommendations.
        Returns a dict with completeness_score (0-100), issues and recommendations.
        """
        validation = {
            'completeness_score': 0,
            'issues': [],
            'recommendations': []
        }

        essential_fields = ['title', 'author', 'slide_count']
        present_fields = sum(1 for field in essential_fields if metadata.get(field))
        validation['completeness_score'] = (present_fields / len(essential_fields)) * 100

        if not metadata.get('title'):
            validation['issues'].append("No document title")
            validation['recommendations'].append("Add a descriptive title to the presentation")

        if not metadata.get('author'):
            validation['issues'].append("No author information")
            validation['recommendations'].append("Set author information in document properties")

        if metadata.get('slide_count', 0) == 0:
            validation['issues'].append("No slides detected")

        if not metadata.get('keywords'):
            validation['recommendations'].append("Add keywords to improve searchability")

        return validation