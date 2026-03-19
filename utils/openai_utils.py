"""
Shared OpenAI helpers for model configuration and response text extraction.
"""
import os
from typing import Any, Optional

from openai import OpenAI


DEFAULT_OPENAI_MODEL = "gpt-5.4"


def get_openai_client(api_key: Optional[str] = None) -> OpenAI:
    """Build an OpenAI client from an explicit key or environment."""
    return OpenAI(api_key=api_key or os.environ.get("OPENAI_API_KEY"))


def get_default_model() -> str:
    """Allow env override while keeping GPT-5.4 as the repo default."""
    return os.environ.get("OPENAI_MODEL", DEFAULT_OPENAI_MODEL)


def extract_output_text(response: Any) -> str:
    """
    Read text from an OpenAI Responses API result.
    Falls back to traversing output blocks if `output_text` is unavailable.
    """
    output_text = getattr(response, "output_text", None)
    if output_text:
        return output_text.strip()

    chunks = []
    for item in getattr(response, "output", []) or []:
        for content in getattr(item, "content", []) or []:
            if getattr(content, "type", "") == "output_text":
                chunks.append(getattr(content, "text", ""))
    return "\n".join(chunk for chunk in chunks if chunk).strip()
