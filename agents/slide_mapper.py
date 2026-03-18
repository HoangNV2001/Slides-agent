"""
Slide Mapper Agent: Maps drafted content to template slide layouts.
Decides which template slide to use for each content slide, generates replacement instructions.
"""
import json
import os
from typing import Optional

import anthropic


def map_content_to_template(
    draft: dict,
    template_analysis: dict,
    user_instructions: str = "",
    api_key: Optional[str] = None,
) -> dict:
    """
    Map each drafted slide to a template slide layout and generate text replacement instructions.

    Returns:
        dict with slide_plan: list of {
            source_slide: str,
            replacements: {old_text: new_text},
            content: slide content dict
        }
    """
    client = anthropic.Anthropic(api_key=api_key or os.environ.get("ANTHROPIC_API_KEY"))

    # Build template info
    template_info = []
    for slide in template_analysis.get("slides", []):
        slide_desc = {
            "index": slide["index"],
            "filename": slide["filename"],
            "has_images": slide.get("has_images", False),
            "has_charts": slide.get("has_charts", False),
            "text_shapes": [
                {"shape_name": ts["shape_name"], "current_text": ts["text"][:200]}
                for ts in slide.get("text_shapes", [])
            ],
        }
        template_info.append(slide_desc)

    system_prompt = """You are a presentation builder that maps content to PowerPoint templates.

Given drafted slide content and a template's slide inventory, you must:
1. For each content slide, choose the BEST matching template slide layout.
2. VARY the layouts - don't use the same template slide for every content slide.
3. Generate text replacements: map each template text placeholder to new content.
4. Consider: title slides map to title layouts, content slides to bullet layouts, data to chart layouts, etc.

Output ONLY valid JSON (no markdown fences):
{
  "slide_plan": [
    {
      "draft_slide_number": 1,
      "source_template_slide": "slide1.xml",
      "source_template_index": 0,
      "layout_reason": "Why this template slide was chosen",
      "text_replacements": {
        "Original Title Text": "New Title Text",
        "Original body text or placeholder": "New body content"
      },
      "notes": "Any special handling needed"
    }
  ],
  "strategy_notes": "Overall mapping strategy explanation"
}

IMPORTANT: The text_replacements keys must EXACTLY match text found in the template slide's text_shapes.
Be precise with matching - use the exact text from the template."""

    user_message = f"""Drafted Slides:
{json.dumps(draft.get("slides", []), indent=2, ensure_ascii=False)}

Template Slide Inventory:
{json.dumps(template_info, indent=2, ensure_ascii=False)}

Additional markitdown text from template:
{template_analysis.get("markitdown_text", "")[:5000]}

User instructions: {user_instructions or "None"}

Map each drafted slide to a template slide and generate replacement instructions."""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4096,
            system=system_prompt,
            messages=[{"role": "user", "content": user_message}],
        )

        response_text = response.content[0].text.strip()
        if response_text.startswith("```"):
            response_text = response_text.split("\n", 1)[1]
            if response_text.endswith("```"):
                response_text = response_text[:-3]
            response_text = response_text.strip()

        return json.loads(response_text)

    except json.JSONDecodeError as e:
        return {
            "slide_plan": [],
            "error": f"JSON parse error: {str(e)}",
            "raw_response": response_text if 'response_text' in dir() else "",
        }
    except Exception as e:
        return {
            "slide_plan": [],
            "error": str(e),
        }
