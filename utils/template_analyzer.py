"""
Template analyzer: extracts structure, text inventory, and slide metadata from PPTX templates.
Uses python-pptx only — no external scripts required.
"""
import os
from typing import Optional
from pptx import Presentation
from pptx.util import Inches, Pt, Emu


def analyze_template(template_path: str) -> dict:
    """
    Full template analysis using python-pptx.
    Returns dict with slides list, total count, text inventory, layout info.
    """
    result = {
        "template_path": template_path,
        "slides": [],
        "total_slides": 0,
        "text_inventory": {},
        "slide_layouts_available": [],
    }

    try:
        prs = Presentation(template_path)

        # Collect available layouts
        for layout in prs.slide_layouts:
            result["slide_layouts_available"].append({
                "name": layout.name,
                "placeholders": [
                    {"idx": ph.placeholder_format.idx, "name": ph.name, "type": str(ph.placeholder_format.type)}
                    for ph in layout.placeholders
                ],
            })

        # Analyze each slide
        for slide_idx, slide in enumerate(prs.slides):
            slide_info = {
                "index": slide_idx,
                "slide_id": slide.slide_id,
                "layout_name": slide.slide_layout.name if slide.slide_layout else "Unknown",
                "text_shapes": [],
                "has_images": False,
                "has_charts": False,
                "has_tables": False,
                "shape_count": len(slide.shapes),
            }

            for shape in slide.shapes:
                # Check shape types
                if shape.shape_type is not None:
                    type_val = int(shape.shape_type)
                    if type_val == 13:  # Picture
                        slide_info["has_images"] = True
                    elif type_val == 3:  # Chart
                        slide_info["has_charts"] = True

                if shape.has_table:
                    slide_info["has_tables"] = True

                if shape.has_text_frame:
                    text_content = []
                    for para in shape.text_frame.paragraphs:
                        para_text = para.text.strip()
                        if para_text:
                            text_content.append(para_text)

                    if text_content:
                        shape_data = {
                            "shape_name": shape.name,
                            "shape_id": shape.shape_id,
                            "text": "\n".join(text_content),
                            "is_placeholder": shape.is_placeholder,
                        }
                        if shape.is_placeholder:
                            shape_data["placeholder_idx"] = shape.placeholder_format.idx
                            shape_data["placeholder_type"] = str(shape.placeholder_format.type)
                        slide_info["text_shapes"].append(shape_data)

            result["slides"].append(slide_info)
            result["text_inventory"][f"slide_{slide_idx}"] = slide_info["text_shapes"]

        result["total_slides"] = len(result["slides"])

    except Exception as e:
        result["error"] = str(e)

    return result


def get_template_summary(analysis: dict) -> str:
    """Generate a human-readable summary of the template."""
    lines = [
        f"Template: {os.path.basename(analysis['template_path'])}",
        f"Total slides: {analysis['total_slides']}",
    ]

    # Available layouts
    layouts = analysis.get("slide_layouts_available", [])
    if layouts:
        lines.append(f"Available layouts: {', '.join(l['name'] for l in layouts)}")

    lines.append("")
    lines.append("Slide inventory:")

    for slide in analysis.get("slides", []):
        idx = slide["index"]
        layout = slide.get("layout_name", "?")
        text_shapes = slide["text_shapes"]
        extras = []
        if slide.get("has_images"):
            extras.append("has images")
        if slide.get("has_charts"):
            extras.append("has charts")
        if slide.get("has_tables"):
            extras.append("has tables")
        extra_str = f" ({', '.join(extras)})" if extras else ""

        lines.append(f"\n  Slide {idx} [layout: {layout}]{extra_str}:")
        for shape in text_shapes:
            preview = shape["text"][:100].replace("\n", " | ")
            ph_info = ""
            if shape.get("is_placeholder"):
                ph_info = f" (placeholder idx={shape.get('placeholder_idx')})"
            lines.append(
                f"    - [{shape['shape_name']}]{ph_info}: \"{preview}\""
                + ("..." if len(shape["text"]) > 100 else "")
            )

    return "\n".join(lines)
