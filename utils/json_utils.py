"""
Shared JSON utilities: robust parsing, Unicode sanitization, common error repair.
Used by all agents that call the Claude API and parse JSON responses.
"""
import json
import re


def sanitize_text(text: str) -> str:
    """
    Replace problematic Unicode chars with ASCII equivalents.
    Safe for JSON string values and API transmission.
    """
    if not text:
        return ""
    replacements = {
        "\u2014": " - ",   # em dash
        "\u2013": " - ",   # en dash
        "\u2018": "'",      # left single quote
        "\u2019": "'",      # right single quote
        "\u201c": '"',      # left double quote
        "\u201d": '"',      # right double quote
        "\u2026": "...",    # ellipsis
        "\u00a0": " ",      # non-breaking space
        "\u200b": "",       # zero-width space
        "\ufeff": "",       # BOM
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def extract_json_block(text: str) -> str:
    """Extract the outermost JSON object { ... } from text that may contain extra content."""
    text = text.strip()

    # Remove markdown fences
    if text.startswith("```"):
        text = re.sub(r'^```[a-zA-Z]*\n?', '', text)
    if text.endswith("```"):
        text = text[:-3].strip()

    # Find outermost { ... } by counting braces, respecting strings
    first_brace = text.find("{")
    if first_brace == -1:
        return text

    depth = 0
    last_brace = -1
    in_string = False
    escape_next = False

    for i in range(first_brace, len(text)):
        c = text[i]
        if escape_next:
            escape_next = False
            continue
        if c == '\\':
            escape_next = True
            continue
        if c == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                last_brace = i
                break

    if last_brace > first_brace:
        return text[first_brace:last_brace + 1]
    return text


def fix_json_string(text: str) -> str:
    """Fix unescaped newlines/tabs inside JSON string values."""
    result = []
    in_string = False
    escape_next = False
    for c in text:
        if escape_next:
            result.append(c)
            escape_next = False
            continue
        if c == '\\':
            result.append(c)
            escape_next = True
            continue
        if c == '"':
            in_string = not in_string
            result.append(c)
            continue
        if in_string and c == '\n':
            result.append('\\n')
            continue
        if in_string and c == '\t':
            result.append('\\t')
            continue
        result.append(c)
    return ''.join(result)


def parse_json_robust(text: str) -> dict:
    """
    Try multiple strategies to parse JSON from an LLM API response.
    Handles: markdown fences, Unicode chars, unescaped newlines, trailing commas.
    Raises json.JSONDecodeError only if ALL strategies fail.
    """
    # Strategy 1: Direct parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Strategy 2: Extract JSON block
    extracted = extract_json_block(text)
    try:
        return json.loads(extracted)
    except json.JSONDecodeError:
        pass

    # Strategy 3: Sanitize Unicode
    sanitized = sanitize_text(extracted)
    try:
        return json.loads(sanitized)
    except json.JSONDecodeError:
        pass

    # Strategy 4: Fix unescaped newlines inside strings
    fixed = fix_json_string(sanitized)
    try:
        return json.loads(fixed)
    except json.JSONDecodeError:
        pass

    # Strategy 5: Remove trailing commas
    no_trailing = re.sub(r',\s*([}\]])', r'\1', fixed)
    try:
        return json.loads(no_trailing)
    except json.JSONDecodeError as e:
        raise json.JSONDecodeError(
            f"All JSON parse strategies failed. Last error: {e.msg}",
            e.doc, e.pos
        )
