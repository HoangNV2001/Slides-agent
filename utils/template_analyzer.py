"""
Template analyzer: extracts structure, text inventory, and slide metadata from PPTX templates.
Uses python-pptx only — no external scripts required.
"""
import os
from collections import Counter
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
        "mapping_slides": [],
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

        slide_width = int(prs.slide_width)
        slide_height = int(prs.slide_height)
        result["slide_width"] = slide_width
        result["slide_height"] = slide_height

        all_texts = []

        # Analyze each slide
        for slide_idx, slide in enumerate(prs.slides):
            slide_info = {
                "index": slide_idx,
                "slide_id": slide.slide_id,
                "layout_name": slide.slide_layout.name if slide.slide_layout else "Unknown",
                "text_shapes": [],
                "image_shapes": [],
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
                        slide_info["image_shapes"].append({
                            "shape_name": shape.name,
                            "shape_id": shape.shape_id,
                            "left": int(shape.left),
                            "top": int(shape.top),
                            "width": int(shape.width),
                            "height": int(shape.height),
                        })
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
                            "left": int(shape.left),
                            "top": int(shape.top),
                            "width": int(shape.width),
                            "height": int(shape.height),
                        }
                        if shape.is_placeholder:
                            shape_data["placeholder_idx"] = shape.placeholder_format.idx
                            shape_data["placeholder_type"] = str(shape.placeholder_format.type)
                        slide_info["text_shapes"].append(shape_data)
                        all_texts.append(_normalize_text(shape_data["text"]))

            result["slides"].append(slide_info)
            result["text_inventory"][f"slide_{slide_idx}"] = slide_info["text_shapes"]

        result["total_slides"] = len(result["slides"])
        repeated_texts = {
            text for text, count in Counter(t for t in all_texts if t).items()
            if count >= max(2, int(result["total_slides"] * 0.4))
        }

        for slide in result["slides"]:
            _annotate_slide_roles(slide, slide_width, slide_height, repeated_texts)

        result["mapping_slides"] = _build_mapping_slide_inventory(result["slides"])

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

    slides_for_mapping = analysis.get("mapping_slides") or analysis.get("slides", [])
    for slide in slides_for_mapping:
        idx = slide["index"]
        layout = slide.get("layout_name", "?")
        simplified = slide.get("simplified_layout", {})
        extras = []
        if slide.get("has_images"):
            extras.append("has images")
        if slide.get("has_charts"):
            extras.append("has charts")
        if slide.get("has_tables"):
            extras.append("has tables")
        extra_str = f" ({', '.join(extras)})" if extras else ""

        lines.append(f"\n  Slide {idx} [layout: {layout}]{extra_str}:")
        lines.append(
            "    - "
            + f"type={simplified.get('layout_kind', 'unknown')}, "
            + f"title_slots={simplified.get('title_slots', 0)}, "
            + f"body_slots={simplified.get('body_slots', 0)}, "
            + f"image_slots={simplified.get('image_slots', 0)}"
        )
        for slot in simplified.get("slots", [])[:6]:
            lines.append(
                f"    - [{slot.get('role')}] {slot.get('shape_name')}: "
                f"\"{slot.get('text_preview', '')}\""
            )

    return "\n".join(lines)


def _normalize_text(text: str) -> str:
    return " ".join((text or "").split()).strip().lower()


def _annotate_slide_roles(slide: dict, slide_width: int, slide_height: int, repeated_texts: set):
    slots = []
    for shape in slide.get("text_shapes", []):
        role = _classify_text_role(shape, slide_width, slide_height, repeated_texts)
        shape["role"] = role
        if role not in {"header", "footer", "decorative"}:
            slots.append({
                "shape_id": shape["shape_id"],
                "shape_name": shape["shape_name"],
                "role": role,
                "text_preview": shape["text"][:80].replace("\n", " | "),
            })

    image_slots = slide.get("image_shapes", [])
    for img in image_slots:
        slots.append({
            "shape_id": img["shape_id"],
            "shape_name": img["shape_name"],
            "role": "image",
            "text_preview": "",
        })

    title_slots = len([slot for slot in slots if slot["role"] == "title"])
    body_slots = len([slot for slot in slots if slot["role"] in {"subtitle", "body", "quote", "comparison"}])
    image_slot_count = len([slot for slot in slots if slot["role"] == "image"])

    if image_slot_count and body_slots >= 1:
        layout_kind = "image_text"
    elif body_slots >= 2:
        layout_kind = "multi_text"
    elif title_slots and body_slots:
        layout_kind = "title_text"
    elif body_slots == 1:
        layout_kind = "single_text"
    else:
        layout_kind = "title_only"

    slide["simplified_layout"] = {
        "layout_kind": layout_kind,
        "title_slots": title_slots,
        "body_slots": body_slots,
        "image_slots": image_slot_count,
        "slots": slots,
    }


def _classify_text_role(shape: dict, slide_width: int, slide_height: int, repeated_texts: set) -> str:
    top = shape.get("top", 0)
    left = shape.get("left", 0)
    width = shape.get("width", 0)
    height = shape.get("height", 0)
    text = shape.get("text", "")
    normalized = _normalize_text(text)
    name = (shape.get("shape_name") or "").lower()
    bottom = top + height

    if normalized in repeated_texts and bottom > slide_height * 0.82:
        return "footer"
    if normalized in repeated_texts and top < slide_height * 0.12:
        return "header"
    if bottom > slide_height * 0.88:
        return "footer"
    if top < slide_height * 0.08 and width < slide_width * 0.35:
        return "header"
    if len(normalized) <= 3:
        return "decorative"
    if shape.get("is_placeholder") and "title" in name:
        return "title"
    if top < slide_height * 0.28 and width > slide_width * 0.35 and len(normalized) <= 120:
        return "title"
    if "quote" in name or len(normalized) > 140:
        return "quote"
    if width < slide_width * 0.42:
        return "comparison"
    if top < slide_height * 0.45 and len(normalized) <= 120:
        return "subtitle"
    return "body"


def _build_mapping_slide_inventory(slides: list) -> list:
    selected = []
    seen_fingerprints = set()
    for slide in slides:
        simplified = slide.get("simplified_layout", {})
        fingerprint = (
            simplified.get("layout_kind"),
            simplified.get("title_slots", 0),
            simplified.get("body_slots", 0),
            simplified.get("image_slots", 0),
            slide.get("has_charts", False),
            slide.get("has_tables", False),
        )
        if fingerprint in seen_fingerprints:
            continue
        seen_fingerprints.add(fingerprint)
        selected.append(slide)
    return selected
