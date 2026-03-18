"""
Document parser: extracts text from PDF, TXT, JSON files.
"""
import json
import os


def parse_document(file_path: str) -> str:
    """Parse a document and return its text content."""
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".txt":
        return _parse_txt(file_path)
    elif ext == ".json":
        return _parse_json(file_path)
    elif ext == ".pdf":
        return _parse_pdf(file_path)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Supported: .pdf, .txt, .json")


def _parse_txt(file_path: str) -> str:
    with open(file_path, "r", encoding="utf-8") as f:
        return f.read()


def _parse_json(file_path: str) -> str:
    with open(file_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return _json_to_text(data)


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
    """Extract text from PDF using pdfplumber (better layout), fallback to pypdf."""
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

    # Fallback to pypdf
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
