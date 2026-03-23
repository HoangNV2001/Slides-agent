"""
Content Drafter Agent: Analyzes source document + template and drafts slide content.
Uses OpenAI to generate structured slide content.
"""
import json
import os
from typing import Optional

try:
    from ..utils.json_utils import sanitize_text as _sanitize_text, parse_json_robust as _parse_json_robust
    from ..utils.openai_utils import (
        extract_output_text as _extract_output_text,
        get_default_model as _get_default_model,
        get_openai_client as _get_openai_client,
    )
except ImportError:
    import sys
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
    from utils.json_utils import sanitize_text as _sanitize_text, parse_json_robust as _parse_json_robust
    from utils.openai_utils import (
        extract_output_text as _extract_output_text,
        get_default_model as _get_default_model,
        get_openai_client as _get_openai_client,
    )


def draft_slide_content(
    document_text: str,
    template_summary: str,
    document_images: Optional[list] = None,
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
    client = _get_openai_client(api_key)

    # Sanitize inputs
    document_text = _sanitize_text(document_text or "")
    template_summary = _sanitize_text(template_summary or "")
    user_instructions = _sanitize_text(user_instructions or "")
    document_images = document_images or []

    system_prompt = """You are a presentation content strategist. Your job is to analyze a source document and create compelling slide content that will be applied to a PowerPoint template.

Rules:
1. Decide an APPROPRIATE number of slides based on the source document and user instructions.
2. Aim for a concise but complete deck. In most cases, choose between 5 and 12 slides unless the material clearly justifies more or fewer.
3. Each slide must have: title, body content, speaker notes, and visual suggestions.
4. Vary slide types: use title slides, content slides, comparison slides, data slides, quote/highlight slides, section dividers.
5. Keep text concise - slides should have bullet points or short phrases, not paragraphs.
6. Optimize for slide fit: prefer 3-5 bullet points per slide, keep each bullet short, and avoid long titles or subtitles.
7. Include [IMAGE: description] or [CHART: description] or [ICON: description] placeholders where visuals would help.
8. The first slide should be a title/cover slide.
9. The last slide can be a summary, CTA, or closing slide.
10. Preserve the original language of the requested content.
11. If the content is in Vietnamese, keep full Vietnamese diacritics exactly as normal natural Vietnamese writing.
12. Avoid typographic punctuation that often breaks JSON formatting, such as smart quotes or em dashes.
13. If relevant source document images are available, assign them to suitable slides using source_image_ids.
14. Only assign an image when its caption or nearby text clearly matches the slide topic. Do not assign images arbitrarily.
15. If a point would make the slide crowded, split it into another slide or compress it into a shorter phrase.

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
      "source_image_ids": ["img_1"],
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
        + "\n---\n\nAvailable Source Images:\n---\n"
        + json.dumps(
            [
                {
                    "id": image.get("id"),
                    "page": image.get("page"),
                    "caption": image.get("caption"),
                    "nearby_text": image.get("nearby_text", "")[:240],
                    "context_keywords": image.get("context_keywords", []),
                }
                for image in document_images[:20]
            ],
            indent=2,
            ensure_ascii=False,
        )
        + "\n---\n\nRequirements:\n"
        + f"- Additional instructions: {user_instructions or 'None'}\n\n"
        + "Please draft the slide content. Remember to output ONLY valid JSON."
    )

    try:
        response = client.responses.create(
            model=_get_default_model(),
            max_output_tokens=4096,
            instructions=system_prompt,
            input=user_message,
        )

        response_text = _extract_output_text(response)

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
    client = _get_openai_client(api_key)

    system_prompt = """You are a presentation content editor. You will receive a current slide draft and user feedback.
Update the draft according to the feedback while maintaining quality and coherence.
Preserve the original language of the draft. If the content is in Vietnamese, keep full Vietnamese diacritics exactly.
Avoid typographic punctuation that often breaks JSON formatting, such as smart quotes or em dashes.
Output ONLY valid JSON with the same structure as the input draft."""

    draft_json = json.dumps(current_draft, indent=2, ensure_ascii=False)
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
        response = client.responses.create(
            model=_get_default_model(),
            max_output_tokens=4096,
            instructions=system_prompt,
            input=user_message,
        )

        response_text = _extract_output_text(response)
        response_text = _sanitize_text(response_text)
        return _parse_json_robust(response_text)

    except Exception as e:
        return {**current_draft, "error": f"Refinement failed: {str(e)}"}


def review_mapped_draft(
    current_draft: dict,
    slide_plan: dict,
    template_analysis: dict,
    document_text: str = "",
    user_feedback: str = "",
    api_key: Optional[str] = None,
) -> dict:
    """
    Refine drafted slide content after mapping so it better fits the selected
    template slide layouts and their visual composition.
    """
    client = _get_openai_client(api_key)

    source_slides = {
        slide.get("index"): slide
        for slide in (template_analysis.get("mapping_slides") or template_analysis.get("slides") or [])
    }

    mapped_context = []
    plan_items = slide_plan.get("slide_plan", []) or []
    for item in plan_items:
        draft_slide_number = item.get("draft_slide_number")
        template_index = item.get("source_slide_index")
        template_slide = source_slides.get(template_index, {})
        mapped_context.append({
            "draft_slide_number": draft_slide_number,
            "source_slide_index": template_index,
            "layout_reason": item.get("layout_reason", ""),
            "visual_layout": (template_slide.get("visual_layout") or {}),
            "simplified_layout": (template_slide.get("simplified_layout") or {}),
            "has_images": template_slide.get("has_images", False),
            "has_charts": template_slide.get("has_charts", False),
            "has_tables": template_slide.get("has_tables", False),
        })

    system_prompt = """You are a presentation review editor working after slide-to-template mapping.
You will receive:
- the current drafted slides
- the selected template slide for each drafted slide
- the true visual layout information for those selected template slides

Your job is to refine the draft so each slide actually fits the mapped layout.

Rules:
1. Keep the same slide count and slide_number values unless the input is malformed.
2. Preserve the original language exactly. If the content is in Vietnamese, keep full Vietnamese diacritics.
3. Rewrite conservatively. Improve fit, clarity, and density without changing the presentation's meaning.
4. Use the mapped visual layout aggressively:
   - sparse hero/title layouts -> shorter titles, fewer body lines
   - two-column or comparison layouts -> parallel phrasing and balanced bullet counts
   - image-led layouts -> keep body concise and let the visual carry more weight
   - chart/data layouts -> short analytical bullets, not prose blocks
5. If a slide seems mismatched to its mapped layout, keep the content concise and add a short layout_review_note explaining the risk.
6. Keep bullet_points and body synchronized. Do not create long paragraphs.
7. Keep visual_suggestion aligned with the mapped layout's actual visual emphasis.
8. template_slide_hint should be updated to reflect the mapped layout actually being used.
9. Avoid typographic punctuation that often breaks JSON formatting, such as smart quotes or em dashes.

Output ONLY valid JSON with this structure:
{
  "outline": "updated overview",
  "slides": [
    {
      "slide_number": 1,
      "slide_type": "title|content|comparison|data|quote|section_divider|closing",
      "title": "Slide title",
      "subtitle": "Optional subtitle",
      "body": "Short body text with \\n separators",
      "bullet_points": ["Point 1", "Point 2"],
      "visual_suggestion": "Short visual direction or null",
      "source_image_ids": ["img_1"],
      "speaker_notes": "Speaker notes",
      "template_slide_hint": "Updated mapped layout hint",
      "layout_review_note": "Short fit note"
    }
  ],
  "review_summary": "Short summary of fit improvements and remaining risks"
}"""

    document_text = _sanitize_text(document_text or "")
    user_feedback = _sanitize_text(user_feedback or "")
    draft_json = json.dumps(current_draft, indent=2, ensure_ascii=False)

    user_message = (
        "Current Draft:\n"
        + draft_json
        + "\n\nMapped Layout Context:\n"
        + json.dumps(mapped_context, indent=2, ensure_ascii=False)
        + "\n\nOriginal Document (reference, truncated):\n"
        + document_text[:6000]
        + "\n\nAdditional reviewer guidance:\n"
        + (user_feedback or "None")
        + "\n\nPlease return the layout-aware revised draft as JSON only."
    )

    try:
        response = client.responses.create(
            model=_get_default_model(),
            max_output_tokens=4096,
            instructions=system_prompt,
            input=user_message,
        )

        response_text = _extract_output_text(response)
        response_text = _sanitize_text(response_text)
        parsed = _parse_json_robust(response_text)
        if "slides" not in parsed:
            parsed["slides"] = current_draft.get("slides", [])
        if "outline" not in parsed:
            parsed["outline"] = current_draft.get("outline", "")
        return parsed

    except Exception as e:
        return {**current_draft, "error": f"Layout-aware review failed: {str(e)}"}
