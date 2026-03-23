"""
Slide Generator Agent: Orchestrates the full slide generation pipeline.
Pure python-pptx — no external scripts. Runs on any OS (macOS, Linux, Windows).
"""
import os
import tempfile
from typing import Optional

try:
    from ..utils.pptx_builder import build_presentation_from_plan
except ImportError:
    import sys
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
    from utils.pptx_builder import build_presentation_from_plan


def generate_slides(
    template_path: str,
    draft: dict,
    slide_plan: dict,
    output_path: str,
    document_images: Optional[list] = None,
    work_dir: str = None,
) -> dict:
    """
    Full slide generation pipeline using pure python-pptx.

    1. Parse the slide_plan from the mapper agent
    2. Convert to builder-friendly format (source_slide_index + text_replacements)
    3. Call build_presentation_from_plan
    4. Return result with status, steps, warnings

    Args:
        template_path: Path to the .pptx template
        draft: The drafted content dict (from content_drafter)
        slide_plan: The mapping plan dict (from slide_mapper)
        output_path: Where to save the final .pptx
        work_dir: Optional working directory (unused in pure python-pptx mode, kept for API compat)
    """
    result = {
        "status": "starting",
        "output_path": output_path,
        "steps": [],
        "warnings": [],
        "review_report": [],
    }

    try:
        document_images = document_images or []
        # Parse the slide plan from the mapper
        plan_items = slide_plan.get("slide_plan", [])
        if not plan_items:
            result["status"] = "error"
            result["error"] = "No slide plan items found"
            return result

        result["steps"].append(f"Slide plan: {len(plan_items)} slides to generate")

        # Convert mapper output to builder format
        builder_plan = []
        draft_slides = draft.get("slides", []) or []
        for item in plan_items:
            draft_slide_number = item.get("draft_slide_number", 0)
            draft_slide = next(
                (slide for slide in draft_slides if slide.get("slide_number") == draft_slide_number),
                draft_slides[draft_slide_number - 1] if 0 < draft_slide_number <= len(draft_slides) else {},
            )
            builder_item = {
                "source_slide_index": item.get("source_slide_index", 0),
                "text_replacements": item.get("text_replacements", {}),
                "draft_content": draft_slide or {},
                "document_images": document_images,
            }
            builder_plan.append(builder_item)
            result["steps"].append(
                f"  Slide {draft_slide_number or '?'}: "
                f"template[{builder_item['source_slide_index']}] "
                f"({item.get('layout_reason', 'no reason')})"
            )

        # Build the presentation
        result["steps"].append("Building presentation...")
        build_result = build_presentation_from_plan(
            template_path=template_path,
            slide_plan=builder_plan,
            output_path=output_path,
        )

        # Merge build results
        result["steps"].extend(build_result.get("steps", []))
        result["warnings"].extend(build_result.get("warnings", []))
        result["review_report"] = build_result.get("review_report", [])
        result["status"] = build_result.get("status", "error")
        result["validation_text"] = build_result.get("validation_text", "")

        if build_result.get("error"):
            result["error"] = build_result["error"]
        if build_result.get("traceback"):
            result["traceback"] = build_result["traceback"]

    except Exception as e:
        result["status"] = "error"
        result["error"] = str(e)
        import traceback
        result["traceback"] = traceback.format_exc()

    return result
