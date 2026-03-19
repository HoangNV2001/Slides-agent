"""
Utility modules for the AI Slide Builder.
Pure Python — no external script dependencies. Runs on any OS.
"""

from .document_parser import parse_document
from .template_analyzer import analyze_template, get_template_summary
from .pptx_builder import (
    build_presentation_from_plan,
    duplicate_slide,
    delete_slide,
    reorder_slides,
    replace_text_in_slide,
    get_slide_text_inventory,
)

__all__ = [
    "parse_document",
    "analyze_template",
    "get_template_summary",
    "build_presentation_from_plan",
    "duplicate_slide",
    "delete_slide",
    "reorder_slides",
    "replace_text_in_slide",
    "get_slide_text_inventory",
]