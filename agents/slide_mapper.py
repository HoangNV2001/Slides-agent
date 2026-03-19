"""
Slide Mapper Agent: Maps drafted content to template slide layouts.
Decides which template slide to use for each content slide, generates replacement instructions.
"""
import json
import os
from typing import Optional

try:
    from ..utils.json_utils import parse_json_robust as _parse_json_robust
    from ..utils.openai_utils import (
        extract_output_text as _extract_output_text,
        get_default_model as _get_default_model,
        get_openai_client as _get_openai_client,
    )
except ImportError:
    import sys
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
    from utils.json_utils import parse_json_robust as _parse_json_robust
    from utils.openai_utils import (
        extract_output_text as _extract_output_text,
        get_default_model as _get_default_model,
        get_openai_client as _get_openai_client,
    )


def map_content_to_template(
    draft: dict,
    template_analysis: dict,
    user_instructions: str = "",
    api_key: Optional[str] = None,
) -> dict:
    """
    Map each drafted slide to a template slide layout and generate text replacement instructions.
    Includes retry logic and robust JSON parsing.
    """
    client = _get_openai_client(api_key)

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

CRITICAL JSON RULES - you MUST follow these exactly:
- Output ONLY a valid JSON object. No other text, no explanations.
- All string values must use double quotes.
- Escape any double quotes inside string values with backslash.
- NEVER put raw newlines inside string values. Use \\n instead.
- No trailing commas after the last item in arrays or objects.
- The text_replacements keys must EXACTLY match text from the template's text_shapes "current_text" field.
- Keep replacement text values SHORT and on a single line.

JSON structure:
{"slide_plan": [{"draft_slide_number": 1, "source_slide_index": 0, "layout_reason": "reason", "text_replacements": {"Original Text": "New text"}}], "strategy_notes": "notes"}"""

    user_message = (
        "Drafted Slides:\n"
        + json.dumps(draft.get("slides", []), indent=2, ensure_ascii=True)
        + "\n\nTemplate Slide Inventory (0-based indices):\n"
        + json.dumps(template_info, indent=2, ensure_ascii=True)
        + "\n\nUser instructions: " + (user_instructions or "None")
        + "\n\nMap each drafted slide to a template slide index and generate replacement instructions."
        + "\nOutput ONLY valid JSON."
    )

    max_retries = 2
    last_error = None
    last_raw = ""

    for attempt in range(max_retries + 1):
        try:
            prompt_input = user_message

            # On retry, send the error as feedback so the model can fix it
            if attempt > 0 and last_error:
                prompt_input = [
                    {"role": "user", "content": [{"type": "input_text", "text": user_message}]},
                    {"role": "assistant", "content": [{"type": "output_text", "text": last_raw[:2000]}]},
                    {"role": "user", "content": [{"type": "input_text", "text": (
                        f"Your JSON had a parse error: {last_error}\n"
                        "Please fix it and output ONLY valid JSON. "
                        "Make sure all strings are properly escaped, "
                        "no trailing commas, no raw newlines inside strings."
                    )}]},
                ]

            response = client.responses.create(
                model=_get_default_model(),
                max_output_tokens=4096,
                instructions=system_prompt,
                input=prompt_input,
            )

            response_text = _extract_output_text(response)
            last_raw = response_text

            parsed = _parse_json_robust(response_text)

            # Validate structure
            if "slide_plan" not in parsed:
                parsed = {"slide_plan": parsed.get("slides", []), "strategy_notes": ""}

            return parsed

        except json.JSONDecodeError as e:
            last_error = str(e)
            if attempt == max_retries:
                return {
                    "slide_plan": [],
                    "error": f"JSON parse error after {max_retries + 1} attempts: {last_error}",
                    "raw_response": last_raw[:3000],
                }
        except Exception as e:
            return {
                "slide_plan": [],
                "error": str(e),
            }

    return {"slide_plan": [], "error": "Unexpected error in mapping"}
