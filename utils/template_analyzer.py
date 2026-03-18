"""
Template analyzer: extracts structure, text inventory, and slide metadata from PPTX templates.
"""
import json
import os
import subprocess
import zipfile
from typing import Optional

from defusedxml import minidom


def analyze_template(template_path: str, work_dir: str = "/tmp/template_analysis") -> dict:
    """
    Full template analysis: structure, text inventory, slide count.
    Returns a dict with all template metadata.
    """
    os.makedirs(work_dir, exist_ok=True)

    result = {
        "template_path": template_path,
        "slides": [],
        "total_slides": 0,
        "text_inventory": {},
        "markitdown_text": "",
    }

    # 1. Extract text with markitdown
    try:
        md_result = subprocess.run(
            ["python", "-m", "markitdown", template_path],
            capture_output=True, text=True, timeout=30
        )
        result["markitdown_text"] = md_result.stdout
    except Exception as e:
        result["markitdown_text"] = f"Error: {e}"

    # 2. Parse PPTX XML structure
    try:
        extract_dir = os.path.join(work_dir, "extracted")
        if os.path.exists(extract_dir):
            subprocess.run(["rm", "-rf", extract_dir])
        os.makedirs(extract_dir, exist_ok=True)

        with zipfile.ZipFile(template_path, 'r') as zf:
            zf.extractall(extract_dir)

        # Parse presentation.xml for slide order
        pres_xml_path = os.path.join(extract_dir, "ppt", "presentation.xml")
        if os.path.exists(pres_xml_path):
            doc = minidom.parse(pres_xml_path)
            sld_id_list = doc.getElementsByTagName("p:sldId") or doc.getElementsByTagName("p:sldIdLst")
            
            # Get slide relationships
            rels_path = os.path.join(extract_dir, "ppt", "_rels", "presentation.xml.rels")
            rels_map = {}
            if os.path.exists(rels_path):
                rels_doc = minidom.parse(rels_path)
                for rel in rels_doc.getElementsByTagName("Relationship"):
                    rid = rel.getAttribute("Id")
                    target = rel.getAttribute("Target")
                    if "slide" in target.lower() and "slideLayout" not in target and "slideMaster" not in target:
                        rels_map[rid] = target

        # 3. Extract text inventory from each slide
        slides_dir = os.path.join(extract_dir, "ppt", "slides")
        if os.path.exists(slides_dir):
            slide_files = sorted(
                [f for f in os.listdir(slides_dir) if f.startswith("slide") and f.endswith(".xml")],
                key=lambda x: int(''.join(filter(str.isdigit, x)) or 0)
            )

            for idx, slide_file in enumerate(slide_files):
                slide_path = os.path.join(slides_dir, slide_file)
                slide_info = _extract_slide_info(slide_path, idx)
                result["slides"].append(slide_info)
                result["text_inventory"][slide_file] = slide_info["text_shapes"]

        result["total_slides"] = len(result["slides"])

    except Exception as e:
        result["error"] = str(e)

    return result


def _extract_slide_info(slide_xml_path: str, index: int) -> dict:
    """Extract text shapes and their content from a slide XML."""
    info = {
        "index": index,
        "filename": os.path.basename(slide_xml_path),
        "text_shapes": [],
        "has_images": False,
        "has_charts": False,
    }

    try:
        doc = minidom.parse(slide_xml_path)

        # Find all shape trees
        sp_elements = doc.getElementsByTagName("p:sp")
        for sp in sp_elements:
            shape_info = _extract_shape_text(sp)
            if shape_info:
                info["text_shapes"].append(shape_info)

        # Check for images
        pic_elements = doc.getElementsByTagName("p:pic")
        if pic_elements:
            info["has_images"] = True

        # Check for charts
        chart_refs = doc.getElementsByTagName("c:chart")
        graphicData = doc.getElementsByTagName("a:graphicData")
        for gd in graphicData:
            uri = gd.getAttribute("uri")
            if "chart" in uri.lower() or "diagram" in uri.lower():
                info["has_charts"] = True
                break

    except Exception as e:
        info["error"] = str(e)

    return info


def _extract_shape_text(sp_element) -> Optional[dict]:
    """Extract text content from a shape element."""
    # Get shape name from nvSpPr
    shape_name = ""
    nvSpPr = sp_element.getElementsByTagName("p:nvSpPr")
    if nvSpPr:
        cNvPr = nvSpPr[0].getElementsByTagName("p:cNvPr") or nvSpPr[0].getElementsByTagName("cNvPr")
        if not cNvPr:
            # Try without namespace
            for child in nvSpPr[0].childNodes:
                if hasattr(child, 'tagName') and 'cNvPr' in child.tagName:
                    shape_name = child.getAttribute("name")
                    break
        else:
            shape_name = cNvPr[0].getAttribute("name")

    # Extract text content
    txBody = sp_element.getElementsByTagName("p:txBody")
    if not txBody:
        return None

    paragraphs = txBody[0].getElementsByTagName("a:p")
    text_content = []
    for para in paragraphs:
        runs = para.getElementsByTagName("a:r")
        para_text = ""
        for run in runs:
            t_elements = run.getElementsByTagName("a:t")
            for t in t_elements:
                if t.firstChild:
                    para_text += t.firstChild.nodeValue or ""
        if para_text.strip():
            text_content.append(para_text.strip())

    if not text_content:
        return None

    return {
        "shape_name": shape_name,
        "text": "\n".join(text_content),
    }


def get_template_summary(analysis: dict) -> str:
    """Generate a human-readable summary of the template."""
    lines = [
        f"Template: {os.path.basename(analysis['template_path'])}",
        f"Total slides: {analysis['total_slides']}",
        "",
        "Slide inventory:",
    ]

    for slide in analysis["slides"]:
        idx = slide["index"]
        fname = slide["filename"]
        text_shapes = slide["text_shapes"]
        extras = []
        if slide.get("has_images"):
            extras.append("has images")
        if slide.get("has_charts"):
            extras.append("has charts")
        extra_str = f" ({', '.join(extras)})" if extras else ""

        lines.append(f"\n  Slide {idx} [{fname}]{extra_str}:")
        for shape in text_shapes:
            preview = shape["text"][:80].replace("\n", " ")
            lines.append(f"    - [{shape['shape_name']}]: \"{preview}...\"" if len(shape["text"]) > 80 else f"    - [{shape['shape_name']}]: \"{preview}\"")

    return "\n".join(lines)
