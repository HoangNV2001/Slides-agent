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
            source_slide_index: int (0-based),
            text_replacements: {old_text: new_text},
        }
    """
    client = anthropic.Anthropic(api_key=api_key or os.environ.get("ANTHROPIC_API_KEY"))

    # Build template info
    template_info = []
    for slide in template_analysis.get("slides", []):
        slide_desc = {
            "index": slide["index"],
            "layout_name": slide.get("layout_name", "Unknown"),
            "has_images": slide.get("has_images", False),
            "has_charts": slide.get("has_charts", False),
            "has_tables": slide.get("has_tables", False),
            "text_shapes": [
                {
                    "shape_name": ts["shape_name"],
                    "current_text": ts["text"][:200],
                    "is_placeholder": ts.get("is_placeholder", False),
                }
                for ts in slide.get("text_shapes", [])
            ],
        }
        template_info.append(slide_desc)

    system_prompt = """You are a presentation builder that maps content to PowerPoint templates.

Given drafted slide content and a template's slide inventory, you must:
1. For each content slide, choose the BEST matching template slide (by its 0-based index).
2. VARY the layouts - don't use the same template slide for every content slide.
3. Generate text replacements: map template text to new content.
4. Title slides map to title layouts, content to bullet layouts, data to chart layouts, etc.

IMPORTANT: The text_replacements keys must EXACTLY match text from the template slide's text_shapes "current_text" field.
Be very precise with the text matching.

Output ONLY valid JSON (no markdown fences):
{
  "slide_plan": [
    {
      "draft_slide_number": 1,
      "source_slide_index": 0,
      "layout_reason": "Why this template slide was chosen",
      "text_replacements": {
        "Exact Original Text": "New replacement text"
      }
    }
  ],
  "strategy_notes": "Overall mapping strategy explanation"
}"""

    user_message = f"""Drafted Slides:
{json.dumps(draft.get("slides", []), indent=2, ensure_ascii=False)}

Template Slide Inventory (0-based indices):
{json.dumps(template_info, indent=2, ensure_ascii=False)}

User instructions: {user_instructions or "None"}

Map each drafted slide to a template slide index and generate replacement instructions."""

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