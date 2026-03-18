"""
Slide Generator Agent: Orchestrates the full slide generation pipeline.
Coordinates template analysis, content mapping, PPTX manipulation, and validation.
"""
import json
import os
import shutil
import subprocess
import re
from typing import Optional

try:
    from ..utils.pptx_builder import PptxBuilder
    from ..utils.template_analyzer import analyze_template
except ImportError:
    import sys, os
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
    from utils.pptx_builder import PptxBuilder
    from utils.template_analyzer import analyze_template


def generate_slides(
    template_path: str,
    draft: dict,
    slide_plan: dict,
    output_path: str,
    work_dir: str = "/tmp/slide_gen",
) -> dict:
    """
    Full slide generation pipeline:
    1. Unpack template
    2. Duplicate slides according to plan
    3. Replace text content
    4. Clean & pack
    5. Validate

    Returns dict with status, output_path, and any warnings.
    """
    result = {
        "status": "starting",
        "output_path": output_path,
        "steps": [],
        "warnings": [],
    }

    scripts_dir = "/mnt/skills/public/pptx/scripts"

    try:
        # Step 1: Unpack
        result["steps"].append("Unpacking template...")
        unpacked_dir = os.path.join(work_dir, "unpacked")
        if os.path.exists(unpacked_dir):
            shutil.rmtree(unpacked_dir)
        subprocess.run(
            ["python", f"{scripts_dir}/office/unpack.py", template_path, unpacked_dir],
            check=True, capture_output=True, text=True
        )
        result["steps"].append("✓ Template unpacked")

        # Step 2: Get current slides and plan duplications
        plan_items = slide_plan.get("slide_plan", [])
        if not plan_items:
            result["status"] = "error"
            result["error"] = "No slide plan items found"
            return result

        # Count how many times each template slide is needed
        slide_usage = {}
        for item in plan_items:
            src = item["source_template_slide"]
            if src not in slide_usage:
                slide_usage[src] = []
            slide_usage[src].append(item)

        result["steps"].append(f"Planning {len(plan_items)} slides from {len(slide_usage)} template layouts...")

        # Step 3: Duplicate slides as needed
        # First use of each template slide doesn't need duplication
        # Subsequent uses need duplication
        slide_file_assignments = {}  # plan_index -> actual slide filename
        duplicated_slides = []

        for src_slide, items in slide_usage.items():
            # First item uses the original
            first_item_idx = items[0]["draft_slide_number"]
            slide_file_assignments[first_item_idx] = src_slide

            # Additional items need duplicates
            for extra_item in items[1:]:
                dup_result = subprocess.run(
                    ["python", f"{scripts_dir}/add_slide.py", unpacked_dir, src_slide],
                    capture_output=True, text=True
                )
                output = dup_result.stdout + dup_result.stderr
                match = re.search(r'(slide\d+\.xml)', output)
                if match:
                    new_slide = match.group(1)
                    slide_file_assignments[extra_item["draft_slide_number"]] = new_slide
                    duplicated_slides.append(new_slide)
                    result["steps"].append(f"  Duplicated {src_slide} → {new_slide}")
                else:
                    result["warnings"].append(f"Could not duplicate {src_slide}: {output[:200]}")
                    # Fallback: use original
                    slide_file_assignments[extra_item["draft_slide_number"]] = src_slide

        result["steps"].append(f"✓ Slide structure ready ({len(duplicated_slides)} duplications)")

        # Step 4: Apply text replacements
        result["steps"].append("Applying text replacements...")
        slides_dir = os.path.join(unpacked_dir, "ppt", "slides")

        for item in plan_items:
            slide_num = item["draft_slide_number"]
            slide_file = slide_file_assignments.get(slide_num)
            if not slide_file:
                result["warnings"].append(f"No slide file for draft slide {slide_num}")
                continue

            replacements = item.get("text_replacements", {})
            if not replacements:
                continue

            slide_path = os.path.join(slides_dir, slide_file)
            if not os.path.exists(slide_path):
                result["warnings"].append(f"Slide file not found: {slide_file}")
                continue

            try:
                with open(slide_path, "r", encoding="utf-8") as f:
                    content = f.read()

                replacements_made = 0
                for old_text, new_text in replacements.items():
                    if old_text in content:
                        # Escape for XML
                        safe_new = (new_text
                                    .replace("&", "&amp;")
                                    .replace("<", "&lt;")
                                    .replace(">", "&gt;"))
                        content = content.replace(old_text, safe_new)
                        replacements_made += 1
                    else:
                        result["warnings"].append(
                            f"Text not found in {slide_file}: '{old_text[:50]}...'"
                        )

                with open(slide_path, "w", encoding="utf-8") as f:
                    f.write(content)

                result["steps"].append(
                    f"  {slide_file}: {replacements_made}/{len(replacements)} replacements"
                )

            except Exception as e:
                result["warnings"].append(f"Error editing {slide_file}: {str(e)}")

        result["steps"].append("✓ Text replacements applied")

        # Step 5: Reorder slides
        result["steps"].append("Reordering slides...")
        ordered_slides = []
        for item in sorted(plan_items, key=lambda x: x["draft_slide_number"]):
            slide_file = slide_file_assignments.get(item["draft_slide_number"])
            if slide_file and slide_file not in ordered_slides:
                ordered_slides.append(slide_file)

        # Update presentation.xml slide order
        _update_slide_order(unpacked_dir, ordered_slides)
        result["steps"].append(f"✓ Slides reordered: {len(ordered_slides)} slides")

        # Step 6: Clean and pack
        result["steps"].append("Cleaning and packing...")
        subprocess.run(
            ["python", f"{scripts_dir}/clean.py", unpacked_dir],
            capture_output=True, text=True
        )
        subprocess.run(
            ["python", f"{scripts_dir}/office/pack.py",
             unpacked_dir, output_path, "--original", template_path],
            capture_output=True, text=True, check=True
        )
        result["steps"].append("✓ PPTX packed")

        # Step 7: Validate
        result["steps"].append("Validating output...")
        try:
            val_result = subprocess.run(
                ["python", "-m", "markitdown", output_path],
                capture_output=True, text=True, timeout=30
            )
            output_text = val_result.stdout
            # Check for leftover placeholders
            import re as re_mod
            placeholders = re_mod.findall(
                r'\b[Xx]{3,}\b|lorem|ipsum|\bTODO\b|\[insert',
                output_text, re_mod.IGNORECASE
            )
            if placeholders:
                result["warnings"].append(
                    f"Found possible leftover placeholders: {placeholders[:5]}"
                )
            result["validation_text"] = output_text[:3000]
        except Exception as e:
            result["warnings"].append(f"Validation error: {str(e)}")

        result["status"] = "success"
        result["steps"].append("✓ Generation complete!")

    except Exception as e:
        result["status"] = "error"
        result["error"] = str(e)
        import traceback
        result["traceback"] = traceback.format_exc()

    return result


def _update_slide_order(unpacked_dir: str, ordered_slides: list):
    """Update the slide order in presentation.xml."""
    from defusedxml import minidom

    pres_path = os.path.join(unpacked_dir, "ppt", "presentation.xml")
    rels_path = os.path.join(unpacked_dir, "ppt", "_rels", "presentation.xml.rels")

    # Build filename -> rId mapping
    filename_to_rid = {}
    if os.path.exists(rels_path):
        rels_doc = minidom.parse(rels_path)
        for rel in rels_doc.getElementsByTagName("Relationship"):
            rid = rel.getAttribute("Id")
            target = rel.getAttribute("Target")
            fname = os.path.basename(target)
            if fname.startswith("slide") and not fname.startswith("slideLayout") and not fname.startswith("slideMaster"):
                filename_to_rid[fname] = rid

    # Parse presentation.xml
    doc = minidom.parse(pres_path)
    sld_id_lst_nodes = doc.getElementsByTagName("p:sldIdLst")
    if not sld_id_lst_nodes:
        return

    sld_id_lst = sld_id_lst_nodes[0]

    # Collect existing sldId elements by rId
    existing = {}
    for sld_id in sld_id_lst.getElementsByTagName("p:sldId"):
        rid = sld_id.getAttribute("r:id")
        existing[rid] = sld_id.cloneNode(True)

    # Clear list
    while sld_id_lst.firstChild:
        sld_id_lst.removeChild(sld_id_lst.firstChild)

    # Re-add in order
    added = set()
    for fname in ordered_slides:
        rid = filename_to_rid.get(fname)
        if rid and rid in existing and rid not in added:
            sld_id_lst.appendChild(existing[rid])
            added.add(rid)

    # Add any remaining slides not in our order (to avoid losing them)
    for rid, node in existing.items():
        if rid not in added:
            sld_id_lst.appendChild(node)

    with open(pres_path, "w", encoding="utf-8") as f:
        doc.writexml(f)
