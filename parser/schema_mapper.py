"""Map parsed workbook data to schema via Gemini. Entity detection is domain-agnostic."""
import json
import os
import re
from typing import Any

from google import genai
from google.genai import types

# Ensure .env is loaded for GEMINI_API_KEY
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

SYSTEM_PROMPT = """You are a data schema expert. Analyze spreadsheet data and identify what kind of data it contains. Work with any domain - financial, HR, sales, inventory, medical, operations, or anything else. Return ONLY valid JSON, no explanation, no markdown."""

USER_PROMPT_TEMPLATE = """Sheet name: {sheet_name}
Columns: {columns}
Sample data: {column_samples}

Analyze this sheet and return ONLY valid JSON:
{{
  "entity": "short description of what this sheet contains, e.g. income_statement, sales_pipeline, employee_roster, inventory_levels, patient_records. Infer from the data - do not use a fixed list.",
  "time_structure": "one of: annual, monthly, quarterly, mixed, none",
  "mapped_fields": {{
    "original_column_name": "canonical_name"
  }},
  "key_metrics": ["3-5 most important fields"],
  "confidence": 0.0,
  "needs_review": true
}}
Set needs_review to true if confidence < 0.75."""


def _build_column_samples(columns: list[str], sample_rows: list[dict[str, Any]], max_rows: int = 10, max_values_per_col: int = 5) -> str:
    """Build compact evidence: column names + first 5 unique non-empty values per column from up to max_rows."""
    rows = sample_rows[:max_rows]
    lines = []
    for col in columns:
        seen: set[str] = set()
        samples: list[str] = []
        for row in rows:
            val = row.get(col)
            if val is None or val == "":
                continue
            s = str(val).strip()
            if s and s not in seen:
                seen.add(s)
                samples.append(s)
            if len(samples) >= max_values_per_col:
                break
        lines.append(f"  {col}: {samples}")
    return "\n".join(lines) if lines else "(no data)"


def _extract_json(text: str) -> str:
    """Extract JSON from response, stripping markdown code blocks if present."""
    text = text.strip()
    match = re.search(r"```(?:json)?\s*([\s\S]*?)```", text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return text


def map_schema(
    sheet_name: str,
    columns: list[str],
    sample_rows: list[dict[str, Any]],
    workbook_id: str,
) -> dict[str, Any]:
    """
    Map sheet data to schema via Gemini. Entity detection is domain-agnostic.
    Uses up to 10 sample rows; builds evidence from first 5 unique non-empty values per column.
    """
    column_samples = _build_column_samples(columns, sample_rows)

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key or not api_key.strip():
        return {
            "entity": "unknown",
            "confidence": 0,
            "error": "GEMINI_API_KEY not configured",
        }

    try:
        client = genai.Client(api_key=api_key.strip())
        user_prompt = USER_PROMPT_TEMPLATE.format(
            sheet_name=sheet_name,
            columns=json.dumps(columns),
            column_samples=column_samples,
        )
        prompt = SYSTEM_PROMPT + "\n\n" + user_prompt

        response = client.models.generate_content(
            model="gemini-2.5-flash-lite",
            contents=prompt,
            config=types.GenerateContentConfig(max_output_tokens=1000),
        )
        text = response.text
        if not text:
            return {"entity": "unknown", "confidence": 0, "error": "Empty Gemini response"}

        json_str = _extract_json(text)
        result = json.loads(json_str)
        if not isinstance(result, dict):
            return {"entity": "unknown", "confidence": 0, "error": "Response was not a JSON object"}

        confidence = result.get("confidence")
        if isinstance(confidence, (int, float)) and confidence < 0.75:
            result["needs_review"] = True

        return result

    except json.JSONDecodeError as e:
        return {"entity": "unknown", "confidence": 0, "error": f"Invalid JSON: {e}"}
    except Exception as e:
        return {"entity": "unknown", "confidence": 0, "error": str(e)}
