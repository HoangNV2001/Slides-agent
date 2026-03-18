"""
Utility modules for the AI Slide Builder.

- document_parser: Extract text from PDF, TXT, JSON files
- template_analyzer: Analyze PPTX template structure and text inventory
- pptx_builder: Low-level PPTX manipulation (duplicate, reorder, replace text)
"""

from .document_parser import parse_document
from .template_analyzer import analyze_template, get_template_summary
from .pptx_builder import PptxBuilder

__all__ = [
    "parse_document",
    "analyze_template",
    "get_template_summary",
    "PptxBuilder",
]
