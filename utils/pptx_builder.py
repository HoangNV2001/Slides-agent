"""
PPTX Builder: slide duplication, reordering, deletion, text replacement.
Pure python-pptx + lxml. No external scripts. Runs on any OS.
"""
import copy
import os
from typing import Dict, List, Optional

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree


def duplicate_slide(prs: Presentation, source_index: int) -> int:
    """
    Duplicate a slide within a presentation.
    Returns the 0-based index of the new slide (appended at end).

    This works at the lxml/OPC level:
      1. Copy the slide XML element tree
      2. Create a new slide part with a new relationship
      3. Copy all relationships from the source slide (images, layouts, etc.)
      4. Append a new <p:sldId> to the presentation's sldIdLst
    """
    source_slide = prs.slides[source_index]
    slide_layout = source_slide.slide_layout

    # Use the internal API to add a blank slide with the same layout
    new_slide = prs.slides.add_slide(slide_layout)

    # Clear everything from the new slide's spTree
    new_sp_tree = new_slide.shapes._spTree
    for child in list(new_sp_tree):
        tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
        # Keep the cNvGrpSpPr and grpSpPr, remove shapes
        if tag not in ("nvGrpSpPr", "grpSpPr"):
            new_sp_tree.remove(child)

    # Deep-copy all shape elements from source
    source_sp_tree = source_slide.shapes._spTree
    for child in source_sp_tree:
        tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
        if tag not in ("nvGrpSpPr", "grpSpPr"):
            new_sp_tree.append(copy.deepcopy(child))

    # Copy relationships (images, charts, etc.) from source to new slide
    # Build mapping: old rId -> new rId
    rid_map = {}
    for rel in source_slide.part.rels.values():
        # Skip the slide layout relationship (already set)
        if rel.reltype == RT.SLIDE_LAYOUT:
            continue
        try:
            if rel.is_external:
                new_rel = new_slide.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
            else:
                new_rel = new_slide.part.rels.get_or_add(rel.reltype, rel.target_part)
            rid_map[rel.rId] = new_rel if isinstance(new_rel, str) else new_rel
        except Exception:
            # Some relationship types may not copy cleanly; skip gracefully
            pass

    # Update rId references in the new slide XML
    if rid_map:
        _update_rids_in_xml(new_sp_tree, rid_map)

    return len(prs.slides) - 1


def _update_rids_in_xml(element, rid_map: dict):
    """Replace old rId references with new ones in an XML element tree."""
    for attr_name, attr_value in element.attrib.items():
        if attr_value in rid_map:
            element.attrib[attr_name] = rid_map[attr_value]
    for child in element:
        _update_rids_in_xml(child, rid_map)


def delete_slide(prs: Presentation, slide_index: int):
    """
    Delete a slide by index from the presentation.
    Works at the lxml/OPC level.
    """
    slide = prs.slides[slide_index]
    rId = None

    # Find the relationship ID for this slide
    for rel_key, rel in prs.part.rels.items():
        if rel.target_part == slide.part:
            rId = rel.rId
            break

    if rId is None:
        raise ValueError(f"Could not find relationship for slide index {slide_index}")

    # Remove from sldIdLst
    nsmap = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    sldIdLst = prs.part._element.find('.//p:sldIdLst', nsmap)
    if sldIdLst is not None:
        for sldId in list(sldIdLst):
            if sldId.get(etree.QName(r_ns, "id")) == rId:
                sldIdLst.remove(sldId)
                break

    # Remove the relationship
    prs.part.drop_rel(rId)


def reorder_slides(prs: Presentation, new_order: List[int]):
    """
    Reorder slides. new_order is a list of current 0-based indices in desired order.
    E.g. [2, 0, 1] means: current slide 2 becomes first, then slide 0, then slide 1.
    """
    nsmap = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
    sldIdLst = prs.part._element.find('.//p:sldIdLst', nsmap)
    if sldIdLst is None:
        raise ValueError("Could not find sldIdLst in presentation XML")

    sld_ids = list(sldIdLst)

    # Validate
    if sorted(new_order) != list(range(len(sld_ids))):
        raise ValueError(f"new_order must be a permutation of 0..{len(sld_ids)-1}, got {new_order}")

    # Remove all
    for sld_id in sld_ids:
        sldIdLst.remove(sld_id)

    # Re-add in new order
    for idx in new_order:
        sldIdLst.append(sld_ids[idx])


def replace_text_in_slide(slide, replacements: Dict[str, str]):
    """
    Replace text in all shapes of a slide.
    replacements: {old_text: new_text}

    Handles text split across multiple runs by working at the paragraph level.
    """
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            _replace_in_paragraph(paragraph, replacements)


def _replace_in_paragraph(paragraph, replacements: Dict[str, str]):
    """
    Replace text in a paragraph, handling text split across runs.

    Strategy:
    1. Concatenate all run texts to get full paragraph text.
    2. If a replacement matches, redistribute the new text across runs,
       preserving the formatting of the first matching run.
    """
    full_text = "".join(run.text for run in paragraph.runs)
    if not full_text:
        return

    new_full_text = full_text
    changed = False
    for old_text, new_text in replacements.items():
        if old_text in new_full_text:
            new_full_text = new_full_text.replace(old_text, new_text)
            changed = True

    if not changed:
        return

    # Redistribute: put all text in first run, clear others
    runs = paragraph.runs
    if not runs:
        return

    runs[0].text = new_full_text
    for run in runs[1:]:
        run.text = ""


def replace_all_text_by_shape_name(slide, shape_name: str, new_text: str):
    """Replace ALL text in a shape identified by name."""
    for shape in slide.shapes:
        if shape.name == shape_name and shape.has_text_frame:
            tf = shape.text_frame
            # Keep formatting from first paragraph/run
            if tf.paragraphs:
                first_para = tf.paragraphs[0]
                # Put new text lines into paragraphs
                lines = new_text.split("\n")

                # Set first paragraph
                if first_para.runs:
                    first_para.runs[0].text = lines[0]
                    for r in first_para.runs[1:]:
                        r.text = ""
                else:
                    first_para.text = lines[0]

                # Add remaining lines as new paragraphs
                for line in lines[1:]:
                    new_para = copy.deepcopy(first_para._p)
                    # Clear runs in copied paragraph and set text
                    runs = new_para.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}r")
                    if runs:
                        t_elem = runs[0].find("{http://schemas.openxmlformats.org/drawingml/2006/main}t")
                        if t_elem is not None:
                            t_elem.text = line
                        for r in runs[1:]:
                            t_elem = r.find("{http://schemas.openxmlformats.org/drawingml/2006/main}t")
                            if t_elem is not None:
                                t_elem.text = ""
                    tf._txBody.append(new_para)

            return True
    return False


def get_slide_text_inventory(slide) -> List[dict]:
    """Get all text shapes and their content from a slide."""
    inventory = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text.strip()
            if text:
                info = {
                    "shape_name": shape.name,
                    "shape_id": shape.shape_id,
                    "text": text,
                    "is_placeholder": shape.is_placeholder,
                }
                if shape.is_placeholder:
                    info["placeholder_idx"] = shape.placeholder_format.idx
                inventory.append(info)
    return inventory


def build_presentation_from_plan(
    template_path: str,
    slide_plan: List[dict],
    output_path: str,
) -> dict:
    """
    High-level: build a full presentation from a template + slide plan.

    slide_plan items:
    {
        "source_slide_index": int,     # 0-based index in template
        "text_replacements": {str: str},  # old_text -> new_text
    }

    Returns dict with status, steps, warnings.
    """
    result = {"status": "starting", "steps": [], "warnings": []}

    try:
        prs = Presentation(template_path)
        original_count = len(prs.slides)
        result["steps"].append(f"Loaded template: {original_count} slides")

        # Phase 1: Duplicate slides as needed.
        # We need len(slide_plan) slides total.
        # Track: plan_index -> actual slide index in the growing presentation.
        slide_index_map = {}  # plan_item_index -> slide_index in prs

        # Count usage of each source slide
        source_usage = {}
        for i, item in enumerate(slide_plan):
            src = item["source_slide_index"]
            if src not in source_usage:
                source_usage[src] = []
            source_usage[src].append(i)

        for src_idx, plan_indices in source_usage.items():
            if src_idx >= original_count:
                result["warnings"].append(f"Source slide index {src_idx} out of range (template has {original_count})")
                continue

            # First usage: use the original slide
            slide_index_map[plan_indices[0]] = src_idx

            # Additional usages: duplicate
            for extra_plan_idx in plan_indices[1:]:
                try:
                    new_idx = duplicate_slide(prs, src_idx)
                    slide_index_map[extra_plan_idx] = new_idx
                    result["steps"].append(f"  Duplicated slide {src_idx} → {new_idx}")
                except Exception as e:
                    result["warnings"].append(f"Failed to duplicate slide {src_idx}: {e}")
                    slide_index_map[extra_plan_idx] = src_idx  # fallback

        result["steps"].append(f"✓ Slide structure ready ({len(prs.slides)} total slides)")

        # Phase 2: Apply text replacements
        result["steps"].append("Applying text replacements...")
        for plan_idx, item in enumerate(slide_plan):
            actual_idx = slide_index_map.get(plan_idx)
            if actual_idx is None:
                result["warnings"].append(f"No slide mapped for plan item {plan_idx}")
                continue

            replacements = item.get("text_replacements", {})
            if not replacements:
                continue

            try:
                slide = prs.slides[actual_idx]
                replace_text_in_slide(slide, replacements)
                result["steps"].append(f"  Slide {actual_idx}: {len(replacements)} replacements applied")
            except Exception as e:
                result["warnings"].append(f"Error replacing text in slide {actual_idx}: {e}")

        result["steps"].append("✓ Text replacements applied")

        # Phase 3: Reorder slides to match plan order, and remove unused slides
        result["steps"].append("Reordering slides...")
        desired_order = [slide_index_map[i] for i in range(len(slide_plan)) if i in slide_index_map]

        # We need to keep only the slides in desired_order.
        # To avoid index shifting issues, we reorder first then delete extras.
        all_indices = set(range(len(prs.slides)))
        used_indices = set(desired_order)
        unused_indices = all_indices - used_indices

        # Build full reorder: desired slides first, then unused
        full_order = desired_order + sorted(unused_indices)
        try:
            reorder_slides(prs, full_order)
            result["steps"].append(f"✓ Slides reordered")
        except Exception as e:
            result["warnings"].append(f"Reorder warning: {e}")

        # Delete unused slides (they're now at the end)
        # Delete from the back to avoid index shifting
        num_to_keep = len(desired_order)
        num_total = len(prs.slides)
        for i in range(num_total - 1, num_to_keep - 1, -1):
            try:
                delete_slide(prs, i)
            except Exception as e:
                result["warnings"].append(f"Could not remove unused slide {i}: {e}")

        result["steps"].append(f"✓ Final presentation: {len(prs.slides)} slides")

        # Phase 4: Save
        prs.save(output_path)
        result["steps"].append(f"✓ Saved to {os.path.basename(output_path)}")

        # Phase 5: Validate
        result["steps"].append("Validating output...")
        try:
            val_prs = Presentation(output_path)
            val_text = ""
            for i, slide in enumerate(val_prs.slides):
                slide_text = []
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        slide_text.append(shape.text_frame.text)
                val_text += f"--- Slide {i + 1} ---\n" + "\n".join(slide_text) + "\n\n"

            # Check for leftover placeholders
            import re
            placeholders = re.findall(
                r'\b[Xx]{3,}\b|lorem|ipsum|\bTODO\b|\[insert',
                val_text, re.IGNORECASE
            )
            if placeholders:
                result["warnings"].append(f"Possible leftover placeholders: {placeholders[:5]}")
            result["validation_text"] = val_text[:3000]
        except Exception as e:
            result["warnings"].append(f"Validation error: {e}")

        result["status"] = "success"
        result["steps"].append("✓ Generation complete!")

    except Exception as e:
        result["status"] = "error"
        result["error"] = str(e)
        import traceback
        result["traceback"] = traceback.format_exc()

    return result