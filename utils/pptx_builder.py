"""
PPTX Builder: handles slide duplication, reordering, deletion, text replacement.
Uses python-pptx for reliable manipulation.
"""
import copy
import json
import os
import re
import shutil
import subprocess
import zipfile
from typing import Dict, List, Optional

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN
from defusedxml import minidom


class PptxBuilder:
    """Build presentations from templates by duplicating/reordering slides and replacing content."""

    def __init__(self, template_path: str, work_dir: str = "/tmp/pptx_build"):
        self.template_path = template_path
        self.work_dir = work_dir
        self.unpacked_dir = os.path.join(work_dir, "unpacked")
        self.scripts_dir = "/mnt/skills/public/pptx/scripts"
        os.makedirs(work_dir, exist_ok=True)

    def unpack(self):
        """Unpack the template PPTX for manipulation."""
        if os.path.exists(self.unpacked_dir):
            shutil.rmtree(self.unpacked_dir)
        subprocess.run(
            ["python", f"{self.scripts_dir}/office/unpack.py", self.template_path, self.unpacked_dir],
            check=True, capture_output=True, text=True
        )

    def get_slide_list(self) -> List[str]:
        """Get current slide filenames from presentation.xml."""
        pres_path = os.path.join(self.unpacked_dir, "ppt", "presentation.xml")
        doc = minidom.parse(pres_path)

        # Get relationship mapping
        rels_path = os.path.join(self.unpacked_dir, "ppt", "_rels", "presentation.xml.rels")
        rels_map = {}
        if os.path.exists(rels_path):
            rels_doc = minidom.parse(rels_path)
            for rel in rels_doc.getElementsByTagName("Relationship"):
                rid = rel.getAttribute("Id")
                target = rel.getAttribute("Target")
                rels_map[rid] = target

        slides = []
        sld_id_lst = doc.getElementsByTagName("p:sldIdLst")
        if sld_id_lst:
            for sld_id in sld_id_lst[0].getElementsByTagName("p:sldId"):
                rid = sld_id.getAttribute("r:id")
                target = rels_map.get(rid, "")
                slide_name = os.path.basename(target)
                if slide_name:
                    slides.append(slide_name)

        return slides

    def duplicate_slide(self, source_slide: str) -> str:
        """Duplicate a slide using the add_slide script. Returns the new slide filename."""
        result = subprocess.run(
            ["python", f"{self.scripts_dir}/add_slide.py", self.unpacked_dir, source_slide],
            capture_output=True, text=True, check=True
        )
        # Parse output for new slide info
        output = result.stdout
        # The script outputs the sldId XML element to add
        # Extract new slide filename from output
        new_slide = None
        for line in output.split("\n"):
            if "slide" in line.lower() and ".xml" in line.lower():
                match = re.search(r'(slide\d+\.xml)', line)
                if match:
                    new_slide = match.group(1)
                    break

        return new_slide or output

    def build_slide_sequence(self, slide_plan: List[dict]) -> List[str]:
        """
        Build slide sequence from a plan.
        Each plan item: {"source_slide": "slide1.xml", "content": {...}}
        Returns list of slide filenames in order.
        """
        new_slides = []
        for item in slide_plan:
            source = item["source_slide"]
            # Duplicate the source slide
            dup_result = subprocess.run(
                ["python", f"{self.scripts_dir}/add_slide.py", self.unpacked_dir, source],
                capture_output=True, text=True
            )
            output = dup_result.stdout + dup_result.stderr
            # Find new slide filename
            match = re.search(r'(slide\d+\.xml)', output)
            if match:
                new_slides.append(match.group(1))
            else:
                new_slides.append(f"dup_of_{source}")

        return new_slides

    def replace_text_in_slide(self, slide_filename: str, replacements: Dict[str, str]):
        """
        Replace text in a slide's XML.
        replacements: {old_text: new_text}
        """
        slide_path = os.path.join(self.unpacked_dir, "ppt", "slides", slide_filename)
        if not os.path.exists(slide_path):
            raise FileNotFoundError(f"Slide not found: {slide_path}")

        with open(slide_path, "r", encoding="utf-8") as f:
            content = f.read()

        for old_text, new_text in replacements.items():
            content = content.replace(old_text, _xml_escape(new_text))

        with open(slide_path, "w", encoding="utf-8") as f:
            f.write(content)

    def replace_text_structured(self, slide_filename: str, shape_replacements: List[dict]):
        """
        Replace text in specific shapes, preserving formatting.
        shape_replacements: [{"shape_name": "Title 1", "new_text": "...", "paragraphs": [...]}]
        """
        slide_path = os.path.join(self.unpacked_dir, "ppt", "slides", slide_filename)
        if not os.path.exists(slide_path):
            raise FileNotFoundError(f"Slide not found: {slide_path}")

        doc = minidom.parse(slide_path)
        sp_elements = doc.getElementsByTagName("p:sp")

        for sp in sp_elements:
            shape_name = _get_shape_name(sp)
            for repl in shape_replacements:
                if repl.get("shape_name") == shape_name or _text_matches(sp, repl.get("match_text", "")):
                    _replace_shape_text(sp, repl)

        with open(slide_path, "w", encoding="utf-8") as f:
            doc.writexml(f)

    def update_slide_order(self, slide_filenames: List[str]):
        """Reorder slides in presentation.xml to match the given list."""
        pres_path = os.path.join(self.unpacked_dir, "ppt", "presentation.xml")

        with open(pres_path, "r", encoding="utf-8") as f:
            content = f.read()

        doc = minidom.parseString(content)

        # Build reverse mapping: slide filename -> rId
        rels_path = os.path.join(self.unpacked_dir, "ppt", "_rels", "presentation.xml.rels")
        filename_to_rid = {}
        if os.path.exists(rels_path):
            rels_doc = minidom.parse(rels_path)
            for rel in rels_doc.getElementsByTagName("Relationship"):
                rid = rel.getAttribute("Id")
                target = rel.getAttribute("Target")
                fname = os.path.basename(target)
                filename_to_rid[fname] = rid

        # Get existing sldId elements
        sld_id_lst = doc.getElementsByTagName("p:sldIdLst")[0]
        existing_ids = {}
        for sld_id in sld_id_lst.getElementsByTagName("p:sldId"):
            rid = sld_id.getAttribute("r:id")
            existing_ids[rid] = sld_id.cloneNode(True)

        # Clear and rebuild
        while sld_id_lst.firstChild:
            sld_id_lst.removeChild(sld_id_lst.firstChild)

        for fname in slide_filenames:
            rid = filename_to_rid.get(fname)
            if rid and rid in existing_ids:
                sld_id_lst.appendChild(existing_ids[rid])

        with open(pres_path, "w", encoding="utf-8") as f:
            doc.writexml(f)

    def clean_and_pack(self, output_path: str):
        """Clean orphaned files and pack back into PPTX."""
        # Clean
        subprocess.run(
            ["python", f"{self.scripts_dir}/clean.py", self.unpacked_dir],
            capture_output=True, text=True
        )
        # Pack
        subprocess.run(
            ["python", f"{self.scripts_dir}/office/pack.py",
             self.unpacked_dir, output_path, "--original", self.template_path],
            capture_output=True, text=True, check=True
        )

    def simple_text_replace(self, output_path: str, all_replacements: Dict[str, Dict[str, str]]):
        """
        High-level: unpack, replace text in multiple slides, clean & pack.
        all_replacements: {slide_filename: {old_text: new_text, ...}, ...}
        """
        self.unpack()
        for slide_file, replacements in all_replacements.items():
            self.replace_text_in_slide(slide_file, replacements)
        self.clean_and_pack(output_path)


def _xml_escape(text: str) -> str:
    """Escape text for XML content."""
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    text = text.replace('"', "&quot;")
    text = text.replace("'", "&apos;")
    return text


def _get_shape_name(sp_element) -> str:
    """Extract shape name from a sp element."""
    nvSpPr = sp_element.getElementsByTagName("p:nvSpPr")
    if nvSpPr:
        for child in nvSpPr[0].childNodes:
            if hasattr(child, 'tagName') and 'cNvPr' in getattr(child, 'tagName', ''):
                return child.getAttribute("name")
    return ""


def _text_matches(sp_element, match_text: str) -> bool:
    """Check if a shape contains specific text."""
    if not match_text:
        return False
    txBody = sp_element.getElementsByTagName("p:txBody")
    if not txBody:
        return False
    full_text = ""
    for t in txBody[0].getElementsByTagName("a:t"):
        if t.firstChild:
            full_text += t.firstChild.nodeValue or ""
    return match_text in full_text


def _replace_shape_text(sp_element, replacement: dict):
    """Replace all text in a shape with new content."""
    txBody = sp_element.getElementsByTagName("p:txBody")
    if not txBody:
        return

    new_text = replacement.get("new_text", "")
    paragraphs = txBody[0].getElementsByTagName("a:p")

    if paragraphs:
        # Use first paragraph as template for formatting
        first_para = paragraphs[0]
        # Get run properties from first run
        first_runs = first_para.getElementsByTagName("a:r")
        rpr_template = None
        if first_runs:
            rprs = first_runs[0].getElementsByTagName("a:rPr")
            if rprs:
                rpr_template = rprs[0].cloneNode(True)

        # Remove all existing paragraphs except first
        for para in list(paragraphs)[1:]:
            txBody[0].removeChild(para)

        # Update first paragraph text
        new_lines = new_text.split("\n") if "\n" in new_text else [new_text]

        # Update first paragraph
        _set_paragraph_text(first_para, new_lines[0], rpr_template)

        # Add additional paragraphs
        for line in new_lines[1:]:
            new_para = first_para.cloneNode(True)
            _set_paragraph_text(new_para, line, rpr_template)
            txBody[0].appendChild(new_para)


def _set_paragraph_text(para_element, text: str, rpr_template=None):
    """Set the text of a paragraph element."""
    # Remove existing runs
    for run in list(para_element.getElementsByTagName("a:r")):
        para_element.removeChild(run)

    # Create new run
    doc = para_element.ownerDocument
    run = doc.createElement("a:r")
    if rpr_template:
        run.appendChild(rpr_template.cloneNode(True))
    t_elem = doc.createElement("a:t")
    t_elem.appendChild(doc.createTextNode(text))
    run.appendChild(t_elem)
    para_element.appendChild(run)
