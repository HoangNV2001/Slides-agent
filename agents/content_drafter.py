"""
Content Drafter Agent: Analyzes source document + template and drafts slide content.
Uses Claude API to generate structured slide content.
"""
import json
import os
import re
from typing import Optional

import anthropic

try:
    from ..utils.json_utils import sanitize_text as _sanitize_text, parse_json_robust as _parse_json_robust
except ImportError:
    import sys
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
    from utils.json_utils import sanitize_text as _sanitize_text, parse_json_robust as _parse_json_robust


def draft_slide_content(
    document_text: str,
    template_summary: str,
    num_slides: int,
    user_instructions: str = "",
    api_key: Optional[str] = None,
) -> dict:
    """
    Draft slide content from a document, considering the template structure.

    Returns:
        dict with:
            - slides: list of slide dicts with title, body, notes, visual_suggestion
            - outline: text overview
    """
    client = anthropic.Anthropic(api_key=api_key or os.environ.get("ANTHROPIC_API_KEY"))

    # Sanitize inputs
    document_text = _sanitize_text(document_text or "")
    template_summary = _sanitize_text(template_summary or "")
    user_instructions = _sanitize_text(user_instructions or "")

    system_prompt = """You are a presentation content strategist. Your job is to analyze a source document and create compelling slide content that will be applied to a PowerPoint template.

Rules:
1. Create EXACTLY the number of slides requested.
2. Each slide must have: title, body content, speaker notes, and visual suggestions.
3. Vary slide types: use title slides, content slides, comparison slides, data slides, quote/highlight slides, section dividers.
4. Keep text concise - slides should have bullet points or short phrases, not paragraphs.
5. Include [IMAGE: description] or [CHART: description] or [ICON: description] placeholders where visuals would help.
6. The first slide should be a title/cover slide.
7. The last slide can be a summary, CTA, or closing slide.
8. Use only ASCII-safe characters in your output. Avoid em dashes, smart quotes, or special Unicode characters.

Output ONLY valid JSON (no markdown fences) with this structure:
{
  "outline": "Brief 2-3 sentence overview of the presentation strategy",
  "slides": [
    {
      "slide_number": 1,
      "slide_type": "title|content|comparison|data|quote|section_divider|closing",
      "title": "Slide Title",
      "subtitle": "Optional subtitle or empty string",
      "body": "Main content - use newline for line breaks between bullet points",
      "bullet_points": ["Point 1", "Point 2", "Point 3"],
      "visual_suggestion": "[IMAGE: description] or [CHART: type - description] or null",
      "speaker_notes": "What the presenter should say",
      "template_slide_hint": "Which template slide layout would work best"
    }
  ]
}"""

    user_message = (
        "Source Document Content:\n---\n"
        + document_text[:15000]
        + "\n---\n\nTemplate Structure:\n---\n"
        + template_summary
        + "\n---\n\nRequirements:\n"
        + f"- Number of slides: {num_slides}\n"
        + f"- Additional instructions: {user_instructions or 'None'}\n\n"
        + "Please draft the slide content. Remember to output ONLY valid JSON."
    )

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4096,
            system=system_prompt,
            messages=[{"role": "user", "content": user_message}],
        )

        response_text = response.content[0].text.strip()

        # Sanitize response before parsing
        response_text = _sanitize_text(response_text)

        # Robust JSON parsing
        parsed = _parse_json_robust(response_text)
        return parsed

    except json.JSONDecodeError as e:
        return {
            "outline": "Error parsing AI response",
            "slides": [],
            "raw_response": response_text[:3000] if 'response_text' in dir() else "",
            "error": str(e),
        }
    except Exception as e:
        return {
            "outline": f"API Error: {str(e)}",
            "slides": [],
            "error": str(e),
        }


def refine_draft(
    current_draft: dict,
    user_feedback: str,
    document_text: str,
    api_key: Optional[str] = None,
) -> dict:
    """Refine the draft based on user feedback."""
    client = anthropic.Anthropic(api_key=api_key or os.environ.get("ANTHROPIC_API_KEY"))

    system_prompt = """You are a presentation content editor. You will receive a current slide draft and user feedback.
Update the draft according to the feedback while maintaining quality and coherence.
Use only ASCII-safe characters. Avoid em dashes, smart quotes, or special Unicode.
Output ONLY valid JSON with the same structure as the input draft."""

    draft_json = json.dumps(current_draft, indent=2, ensure_ascii=True)
    document_text = _sanitize_text(document_text or "")
    user_feedback = _sanitize_text(user_feedback or "")

    user_message = (
        "Current Draft:\n"
        + draft_json
        + "\n\nUser Feedback:\n"
        + user_feedback
        + "\n\nOriginal Document (for reference):\n"
        + document_text[:8000]
        + "\n\nPlease provide the updated draft as JSON only."
    )

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4096,
            system=system_prompt,
            messages=[{"role": "user", "content": user_message}],
        )

        response_text = response.content[0].text.strip()
        response_text = _sanitize_text(response_text)
        return _parse_json_robust(response_text)

    except Exception as e:
        return {**current_draft, "error": f"Refinement failed: {str(e)}"}
