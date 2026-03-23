"""
Document parser: extracts text and images from PDF, TXT, JSON files.
Pure Python, no external script dependencies.
"""
import hashlib
import json
import os
from pathlib import Path
from typing import Dict, List, Optional


def parse_document(file_path: str) -> str:
    """Backward-compatible text-only document parse."""
    return parse_document_bundle(file_path).get("text", "")


def parse_document_bundle(file_path: str, asset_dir: Optional[str] = None) -> dict:
    """Parse a document and return text plus extracted image assets."""
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".txt":
        with open(file_path, "r", encoding="utf-8") as f:
            return {"text": f.read(), "images": []}
    elif ext == ".json":
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {"text": _json_to_text(data), "images": []}
    elif ext == ".pdf":
        return _parse_pdf_bundle(file_path, asset_dir=asset_dir)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Supported: .pdf, .txt, .json")


def _json_to_text(data, indent=0) -> str:
    """Recursively convert JSON to readable text."""
    lines = []
    prefix = "  " * indent
    if isinstance(data, dict):
        for key, value in data.items():
            if isinstance(value, (dict, list)):
                lines.append(f"{prefix}{key}:")
                lines.append(_json_to_text(value, indent + 1))
            else:
                lines.append(f"{prefix}{key}: {value}")
    elif isinstance(data, list):
        for i, item in enumerate(data):
            if isinstance(item, (dict, list)):
                lines.append(f"{prefix}Item {i + 1}:")
                lines.append(_json_to_text(item, indent + 1))
            else:
                lines.append(f"{prefix}- {item}")
    else:
        lines.append(f"{prefix}{data}")
    return "\n".join(lines)


def _parse_pdf(file_path: str) -> str:
    """Backward-compatible text-only PDF parse."""
    return _parse_pdf_bundle(file_path).get("text", "")


def _parse_pdf_bundle(file_path: str, asset_dir: Optional[str] = None) -> dict:
    """Extract text and embedded images from PDF."""
    return {
        "text": _extract_pdf_text(file_path),
        "images": _extract_pdf_images(file_path, asset_dir=asset_dir),
    }


def _extract_pdf_text(file_path: str) -> str:
    """Extract text from PDF. Tries pdfplumber first, falls back to pypdf."""
    try:
        import pdfplumber
        text_parts = []
        with pdfplumber.open(file_path) as pdf:
            for i, page in enumerate(pdf.pages):
                page_text = page.extract_text()
                if page_text:
                    text_parts.append(f"--- Page {i + 1} ---\n{page_text}")
        if text_parts:
            return "\n\n".join(text_parts)
    except Exception:
        pass

    try:
        from pypdf import PdfReader
        reader = PdfReader(file_path)
        text_parts = []
        for i, page in enumerate(reader.pages):
            page_text = page.extract_text()
            if page_text:
                text_parts.append(f"--- Page {i + 1} ---\n{page_text}")
        return "\n\n".join(text_parts) if text_parts else "Could not extract text from PDF."
    except Exception as e:
        return f"Error reading PDF: {e}"


def _extract_pdf_images(file_path: str, asset_dir: Optional[str] = None) -> List[Dict[str, object]]:
    """Extract embedded raster images from PDF pages."""
    output_dir = Path(asset_dir or (Path(file_path).parent / f"{Path(file_path).stem}_assets"))
    output_dir.mkdir(parents=True, exist_ok=True)

    images: List[Dict[str, object]] = []
    seen_hashes = set()
    page_image_metadata = _extract_pdf_image_metadata(file_path)

    try:
        from pypdf import PdfReader

        reader = PdfReader(file_path)
        for page_num, page in enumerate(reader.pages, start=1):
            metadata_items = page_image_metadata.get(page_num, [])
            for image_index, page_image in enumerate(getattr(page, "images", []) or []):
                image_bytes = getattr(page_image, "data", b"") or b""
                if not image_bytes:
                    continue

                image_hash = hashlib.sha1(image_bytes).hexdigest()
                if image_hash in seen_hashes:
                    continue
                seen_hashes.add(image_hash)

                image_id = f"img_{len(images) + 1}"
                image_ext = _normalize_image_extension(getattr(page_image, "name", ""))
                image_path = output_dir / f"{image_id}{image_ext}"
                image_path.write_bytes(image_bytes)

                metadata = metadata_items[image_index] if image_index < len(metadata_items) else {}

                images.append({
                    "id": image_id,
                    "path": str(image_path),
                    "page": page_num,
                    "caption": metadata.get("caption") or f"Image extracted from page {page_num}",
                    "nearby_text": metadata.get("nearby_text", ""),
                    "context_keywords": metadata.get("context_keywords", []),
                    "bbox": metadata.get("bbox"),
                })
    except Exception:
        return []

    return images


def _normalize_image_extension(filename: str) -> str:
    suffix = Path(filename or "").suffix.lower()
    if suffix in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"}:
        return suffix
    return ".png"


def _extract_pdf_image_metadata(file_path: str) -> Dict[int, List[Dict[str, object]]]:
    """
    Extract image-position metadata and infer captions from nearby text using pdfplumber.
    Metadata is keyed by 1-based page number and ordered to roughly match page.images.
    """
    try:
        import pdfplumber
    except Exception:
        return {}

    metadata_by_page: Dict[int, List[Dict[str, object]]] = {}
    try:
        with pdfplumber.open(file_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                words = page.extract_words(use_text_flow=True, keep_blank_chars=False) or []
                images = sorted(
                    page.images or [],
                    key=lambda img: (float(img.get("top", 0)), float(img.get("x0", 0))),
                )
                metadata_by_page[page_num] = [
                    _build_image_metadata_for_region(words, image)
                    for image in images
                ]
    except Exception:
        return {}

    return metadata_by_page


def _build_image_metadata_for_region(words: List[dict], image: dict) -> Dict[str, object]:
    x0 = float(image.get("x0", 0))
    x1 = float(image.get("x1", 0))
    top = float(image.get("top", 0))
    bottom = float(image.get("bottom", 0))

    nearby_words = [
        word for word in words
        if _word_near_image(word, x0, x1, top, bottom)
    ]

    nearby_text = _join_words_as_lines(nearby_words)
    caption = _infer_image_caption(words, x0, x1, top, bottom) or nearby_text.split("\n")[0].strip()
    caption = _normalize_caption(caption)

    return {
        "bbox": {"x0": x0, "x1": x1, "top": top, "bottom": bottom},
        "caption": caption,
        "nearby_text": nearby_text[:500],
        "context_keywords": _extract_keywords(" ".join(filter(None, [caption, nearby_text]))),
    }


def _word_near_image(word: dict, x0: float, x1: float, top: float, bottom: float) -> bool:
    wx0 = float(word.get("x0", 0))
    wx1 = float(word.get("x1", 0))
    wtop = float(word.get("top", 0))
    wbottom = float(word.get("bottom", 0))

    horizontal_overlap = not (wx1 < x0 - 40 or wx0 > x1 + 40)
    vertical_near = not (wbottom < top - 80 or wtop > bottom + 80)
    return horizontal_overlap and vertical_near


def _infer_image_caption(words: List[dict], x0: float, x1: float, top: float, bottom: float) -> str:
    below_words = [
        word for word in words
        if float(word.get("top", 0)) >= bottom and float(word.get("top", 0)) <= bottom + 70
        and not (float(word.get("x1", 0)) < x0 - 30 or float(word.get("x0", 0)) > x1 + 30)
    ]
    above_words = [
        word for word in words
        if float(word.get("bottom", 0)) <= top and float(word.get("bottom", 0)) >= top - 50
        and not (float(word.get("x1", 0)) < x0 - 30 or float(word.get("x0", 0)) > x1 + 30)
    ]

    below_text = _join_words_as_lines(below_words)
    above_text = _join_words_as_lines(above_words)
    return below_text.split("\n")[0].strip() or above_text.split("\n")[0].strip()


def _join_words_as_lines(words: List[dict]) -> str:
    if not words:
        return ""
    sorted_words = sorted(words, key=lambda w: (round(float(w.get("top", 0)), 1), float(w.get("x0", 0))))
    lines: List[List[str]] = []
    current_top = None
    for word in sorted_words:
        word_top = round(float(word.get("top", 0)), 1)
        text = str(word.get("text", "")).strip()
        if not text:
            continue
        if current_top is None or abs(word_top - current_top) > 4:
            lines.append([text])
            current_top = word_top
        else:
            lines[-1].append(text)
    return "\n".join(" ".join(line) for line in lines if line)


def _normalize_caption(text: str) -> str:
    text = " ".join((text or "").split()).strip(" -:;,.")
    return text[:180]


def _extract_keywords(text: str) -> List[str]:
    tokens = []
    for token in (text or "").lower().split():
        token = "".join(ch for ch in token if ch.isalnum() or ch in {"_", "-"})
        if len(token) >= 4:
            tokens.append(token)
    seen = set()
    result = []
    for token in tokens:
        if token not in seen:
            seen.add(token)
            result.append(token)
    return result[:12]
