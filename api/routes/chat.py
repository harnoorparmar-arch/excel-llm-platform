"""Chat endpoint for querying workbook data via Gemini."""
import json
import os
import re
import secrets
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import pandas as pd
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from google import genai
from google.genai import types

from parser.forecast_engine import run_forecast
from parser.quality_engine import run_quality_checks
from storage.database import get_quality_report, get_schema_mappings, get_workbook, init_db

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

router = APIRouter(prefix="/chat", tags=["chat"])


def _db_path() -> str:
    path = Path(os.getenv("DATABASE_PATH", "./data/platform.db"))
    path.parent.mkdir(parents=True, exist_ok=True)
    return str(path)


class ChatQueryRequest(BaseModel):
    workbook_id: str
    question: str
    conversation_id: Optional[str] = None


class ChatQueryResponse(BaseModel):
    answer: str
    sheets_used: list[str]
    workbook_id: str
    conversation_id: Optional[str] = None


def get_conversation_history(conversation_id: str, db_path: str, max_turns: int = 10) -> list[dict[str, str]]:
    """
    Loads last N turns of conversation.
    Returns list of {role, content} dicts.
    """
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute("""
            SELECT role, content
            FROM conversations
            WHERE conversation_id = ?
            ORDER BY created_at DESC
            LIMIT ?
        """, (conversation_id, max_turns * 2))

        rows = cursor.fetchall()
        conn.close()

        # Reverse to get chronological order
        history = [
            {"role": r["role"], "content": r["content"]}
            for r in reversed(rows)
        ]

        return history

    except Exception:
        return []


def save_conversation_turn(
    conversation_id: str,
    workbook_id: str,
    question: str,
    answer: str,
    db_path: str,
) -> None:
    """
    Saves one Q&A turn to conversation history.
    """
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        now = datetime.now().isoformat()

        # Save user message
        cursor.execute("""
            INSERT INTO conversations
            (conversation_id, workbook_id,
             role, content, created_at)
            VALUES (?, ?, ?, ?, ?)
        """, (conversation_id, workbook_id, "user", question, now))

        # Save assistant message
        cursor.execute("""
            INSERT INTO conversations
            (conversation_id, workbook_id,
             role, content, created_at)
            VALUES (?, ?, ?, ?, ?)
        """, (conversation_id, workbook_id, "assistant", answer, now))

        conn.commit()
        conn.close()

    except Exception as e:
        print(f"Failed to save conversation: {e}")


SYSTEM_PROMPT = """You are a data analyst assistant.
You help users understand, analyze, and
reason about data from spreadsheet files.

You have two modes depending on the question:

MODE 1 - FACT FINDING:
For questions asking about specific values
that exist in the data:
  Answer directly from the evidence provided.
  Cite which sheet the answer comes from.
  Never invent numbers not in the evidence.

MODE 2 - INFERENCE AND REASONING:
For questions asking you to predict, estimate,
calculate ratios, compare, or reason beyond
what is explicitly in the data:
  Use the data as your foundation.
  Apply logical reasoning and domain knowledge.
  You may calculate derived metrics such as:
    - Ratios between two values in the data
    - Percentage changes between periods
    - Comparisons between metrics
    - Predictions based on visible trends
    - What a metric implies about business health
  Always show your reasoning step by step.
  Always label inferred answers clearly as
  estimates or calculations not direct facts.
  Always state what data you used as the basis.

REASONING PRINCIPLE:
When asked about something not explicitly in
the data, always attempt to answer using
proxy variables or related fields.

Before saying data is unavailable:
1. Look for fields that could serve as a
   proxy for what was asked
2. State your assumption clearly
3. Calculate from available data
4. Caveat the answer appropriately

Example pattern:
  User asks about X
  X is not directly in data
  But field Y is related to X
  → Calculate X from Y
  → Say: 'Using Y as a proxy for X...'
  → Give the answer
  → Note the limitation

Only refuse if there is truly no related
data anywhere in the dataset.
Never refuse without first attempting
to reason from proxy fields.

FOR ALL QUESTIONS:
  If human confirmed context is provided use it.
  Prefer the highest confidence sheet when
  multiple sheets have the same metric.
  If you genuinely cannot answer even after
  attempting to reason from related fields,
  explain specifically what data would be needed.
  Be concise and specific.
  Never make up numbers that have no basis
  in the evidence.

If a CONVERSATION HISTORY is provided
use it to understand context for the
current question.

For follow-up questions like:
  'what about year 4?'
  'and the expenses?'
  'how does that compare to year 3?'
  'what was that again?'

Use the conversation history to understand
what the user is referring to.

Always answer the CURRENT QUESTION.
Use history only for context not as
the primary answer source."""


def score_sheet_for_question(
    question_words, sheet_name, entity,
    key_metrics, columns, row_labels
):
    score = 0

    for word in question_words:
        if len(word) <= 2:
            continue

        word_lower = word.lower()

        # Key metrics = highest weight
        # These are AI-identified important fields
        # that specifically describe sheet content
        for metric in key_metrics:
            if word_lower in metric.lower():
                score += 10
                break

        # Entity type = high weight
        # AI-identified description of sheet purpose
        if word_lower in entity.lower():
            score += 7

        # Sheet name = medium weight
        # Just a label, can be misleading
        # e.g. sheet named "REVENUE" might not
        # be the best source for revenue questions
        if word_lower in sheet_name.lower():
            score += 4

        # Column names = low weight
        for col in columns:
            if word_lower in str(col).lower():
                score += 2
                break

        # Row labels = lowest weight
        # Data values not structural identifiers
        for label in row_labels:
            if word_lower in str(label).lower():
                score += 1
                break

    return score


def _clean_sheet_with_pandas(sample_rows: list[dict[str, Any]]) -> tuple[list[str], list[dict[str, Any]]] | None:
    """
    Clean sheet data using pandas. Returns (columns, rows) or None if < 2 rows.
    """
    if not sample_rows:
        return None

    df = pd.DataFrame(sample_rows)

    # Replace empty strings with NaN
    df = df.replace("", pd.NA)

    # Remove rows where all values are NaN
    df = df.dropna(how="all")

    # Remove rows where more than 70% are NaN
    threshold = len(df.columns) * 0.3
    df = df.dropna(thresh=int(threshold))

    # Convert numeric strings to numbers
    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col])
        except (ValueError, TypeError):
            pass

    # Handle single letter columns
    for col in list(df.columns):
        if len(str(col).strip()) == 1:
            non_empty = df[col].notna().sum()
            if len(df) > 0 and non_empty / len(df) > 0.2:
                df = df.rename(columns={col: "label"})
            else:
                df = df.drop(columns=[col])

    # Keep only first column named "label" if there are duplicates
    label_indices = [i for i, c in enumerate(df.columns) if str(c).startswith("label")]
    if len(label_indices) > 1:
        keep_positions = [i for i in range(len(df.columns)) if i not in label_indices[1:]]
        df = df.iloc[:, keep_positions]

    if len(df.columns) == 0:
        return None

    # Remove subtotal and total rows
    # to prevent double counting
    SUBTOTAL_KEYWORDS = {
        'total', 'subtotal', 'sub-total',
        'sub total', 'grand total',
        'sum', 'aggregate', 'net total',
        'overall', 'combined', 'totals'
    }

    # Find label column
    label_col = None
    for col in df.columns:
        if str(col).lower() in ['label', 'a',
                                'description',
                                'item', 'name']:
            label_col = col
            break

    if label_col and label_col in df.columns:
        def is_subtotal_row(val):
            val_str = str(val).strip().lower()
            return any(
                kw in val_str
                for kw in SUBTOTAL_KEYWORDS
            )

        subtotal_mask = df[label_col].apply(
            is_subtotal_row
        )

        # Count removed rows for logging
        subtotal_count = subtotal_mask.sum()
        if subtotal_count > 0:
            print(f"Removed {subtotal_count} "
                  f"subtotal rows from evidence")

        # Keep only non-subtotal rows
        df = df[~subtotal_mask]

    # Cap at 20 rows
    df = df.head(20)

    # Convert to clean rows dropping NaN
    def _to_native(v: Any) -> Any:
        if hasattr(v, "item"):
            return v.item()
        return v

    rows: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        clean_row = {k: _to_native(v) for k, v in row.items() if pd.notna(v)}
        if len(clean_row) >= 2:
            rows.append(clean_row)

    if len(rows) < 2:
        return None

    return (list(df.columns), rows)


def _build_evidence_and_send(
    workbook_id: str,
    file_name: str,
    workbook_sheets: list[dict[str, Any]],
    schema_mappings: list[dict[str, Any]],
    quality_issues: list[dict[str, Any]],
    question: str,
    conversation_history: list[dict[str, str]] | None = None,
) -> tuple[str, list[str]]:
    """
    Steps 1-6: Load, clean, score, select, build evidence, return (answer, sheets_used).
    The Gemini call is done here; returns (answer, sheets_used).
    """
    # STEP 1 - Load all sheets with their data
    sheets_by_name = {s["sheet_name"]: s for s in workbook_sheets}
    loaded: list[dict[str, Any]] = []
    for mapping in schema_mappings:
        sheet_name = mapping["sheet_name"]
        sheet_data = sheets_by_name.get(sheet_name)
        if not sheet_data or not sheet_data.get("sample_rows"):
            continue
        conf = mapping.get("confidence")
        try:
            confidence = float(conf) if conf is not None else 0.0
        except (TypeError, ValueError):
            confidence = 0.0
        loaded.append({
            "sheet_name": sheet_name,
            "entity": mapping.get("entity", "unknown"),
            "key_metrics": mapping.get("key_metrics") or [],
            "columns": sheet_data.get("columns") or [],
            "sample_rows": sheet_data.get("sample_rows", []),
            "confidence": confidence,
        })

    # STEP 2 - Clean each sheet's data with pandas
    cleaned_sheets: list[dict[str, Any]] = []
    for item in loaded:
        result = _clean_sheet_with_pandas(item["sample_rows"])
        if result is None:
            continue
        cols, rows = result
        cleaned_sheets.append({
            "sheet_name": item["sheet_name"],
            "entity": item["entity"],
            "key_metrics": item["key_metrics"],
            "columns": item["columns"],
            "clean_columns": cols,
            "clean_rows": rows,
            "confidence": item["confidence"],
        })

    # STEP 3 & 4 - Data-driven sheet selection: 1 anchor + 2 most relevant (no hardcoded terms)
    if len(cleaned_sheets) <= 3:
        selected = cleaned_sheets[:3]
    else:
        # STEP 1: Score every sheet against the question (metadata + row label data)
        question_words = set(question.lower().split())
        for item in cleaned_sheets:
            row_labels = []
            for row in item.get("clean_rows", []):
                label_val = row.get("label", "")
                if label_val and isinstance(label_val, str):
                    row_labels.append(str(label_val))
            item["score"] = score_sheet_for_question(
                question_words,
                item["sheet_name"],
                item["entity"],
                item["key_metrics"] or [],
                item.get("clean_columns") or [],
                row_labels,
            )

        for item in cleaned_sheets:
            print(
                f"SHEET: {item['sheet_name']} | "
                f"ENTITY: {item['entity']} | "
                f"SCORE: {item['score']} | "
                f"KEY_METRICS: {item['key_metrics'][:3]}"
            )

        # STEP 2: Anchor = highest scoring sheet (tie-break: higher confidence)
        sorted_sheets = sorted(
            cleaned_sheets,
            key=lambda x: (
                x.get("score", 0),
                x.get("confidence", 0),
            ),
            reverse=True,
        )
        anchor = sorted_sheets[0]
        remaining = [s for s in sorted_sheets[1:] if s["sheet_name"] != anchor["sheet_name"]]

        # STEP 3: Take 2 more from next highest scoring
        top_relevant = remaining[:2]

        # STEP 4: Final = 1 anchor + 2 most relevant = 3 sheets
        selected = [anchor] + top_relevant

    # DEBUG: Print sheet scores and selection
    sheet_scores = {item["sheet_name"]: item.get("score", -1) for item in cleaned_sheets}
    print("=== SHEET SCORES ===")
    for sheet_name, score in sorted(
        sheet_scores.items(),
        key=lambda x: x[1],
        reverse=True
    ):
        print(f"  {sheet_name}: {score}")
    print("=== TOP 3 SELECTED ===")
    for sheet in selected:
        print(f"  {sheet.get('sheet_name', 'unknown')}")
    print("====================")

    # Load confirmed connections and decisions from SQLite
    init_db()
    db_path = _db_path()
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT sheet_1, identifier_1,
               sheet_2, identifier_2,
               connection_type
        FROM sheet_connections
        WHERE workbook_id = ? AND confirmed = 1
    """, (workbook_id,))
    connections = cursor.fetchall()

    cursor.execute("""
        SELECT type, decision, notes
        FROM review_decisions
        WHERE workbook_id = ?
    """, (workbook_id,))
    decisions = cursor.fetchall()

    conn.close()

    # Build connections context string
    connections_context = ""

    for conn_row in connections:
        sheet_1, id_1, sheet_2, id_2, conn_type = conn_row

        if conn_type == "same_metric":
            connections_context += (
                f"- '{id_1}' in {sheet_1} sheet is the "
                f"same metric as '{id_2}' in {sheet_2} "
                f"sheet (confirmed by human)\n"
            )
        elif conn_type == "column_match":
            connections_context += (
                f"- Column '{id_1}' in {sheet_1} is the "
                f"same field as column '{id_2}' in "
                f"{sheet_2} (confirmed by human)\n"
            )

    for dec_row in decisions:
        dec_type, decision, notes = dec_row

        if dec_type == "date_format":
            try:
                details = json.loads(notes) if notes else {}
                sheet = details.get("sheet", "")
                col = details.get("column", "")
                connections_context += (
                    f"- Dates in {sheet}.{col} "
                    f"use {decision} format "
                    f"(confirmed by human)\n"
                )
            except Exception:
                pass

        elif dec_type == "unit_scale":
            try:
                details = json.loads(notes) if notes else {}
                sheet = details.get("sheet", "")
                col = details.get("column", "")
                connections_context += (
                    f"- Values in {sheet}.{col} "
                    f"are in {decision} "
                    f"(confirmed by human)\n"
                )
            except Exception:
                pass

        elif dec_type == "low_confidence_sheet":
            if decision.lower().startswith("yes"):
                try:
                    details = json.loads(notes) if notes else {}
                    sheet = details.get("sheet", "")
                    entity = details.get("ai_guess", "")
                    connections_context += (
                        f"- {sheet} sheet confirmed "
                        f"to contain {entity} data "
                        f"(confirmed by human)\n"
                    )
                except Exception:
                    pass

    # STEP 5 - Build evidence JSON
    high_issues = [i for i in quality_issues if i.get("severity") == "high"]
    quality_issues_list = [
        {"severity": i.get("severity", "high"), "sheet": i.get("sheet", "?"), "message": i.get("message", "")}
        for i in high_issues
    ]

    evidence = {
        "file": file_name,
        "sheets": [
            {
                "name": s["sheet_name"],
                "type": s["entity"],
                "columns": s["clean_columns"],
                "rows": s["clean_rows"],
            }
            for s in selected
        ],
        "quality_issues": quality_issues_list,
    }

    if connections_context:
        evidence["human_confirmed_context"] = (
            "The following relationships and facts "
            "have been verified by a human analyst:\n"
            + connections_context
        )

    # Forecast intent detection (two-step: triggers + period-beyond-data)
    sheets_data = [
        {"sheet_name": s["sheet_name"], "columns": s["clean_columns"], "rows": s["clean_rows"]}
        for s in selected
    ]
    question_lower = question.lower()

    # STEP 1: Explicit forecast trigger phrases
    forecast_triggers = [
        "what will", "how much will",
        "will be", "will it be",
        "predict", "forecast", "project",
        "future", "estimate", "extrapolate",
        "going forward", "next period",
        "next year", "next month",
        "next quarter", "trend",
        "growth rate", "by when",
        "when will",
    ]
    is_forecast = any(trigger in question_lower for trigger in forecast_triggers)

    # STEP 2: Question asks about period beyond what exists in data
    if not is_forecast:
        period_patterns = [
            r"year\s*(\d+)",
            r"month\s*(\d+)",
            r"quarter\s*(\d+)",
            r"q(\d+)",
            r"period\s*(\d+)",
            r"week\s*(\d+)",
        ]
        question_periods = []
        for pattern in period_patterns:
            matches = re.findall(pattern, question_lower)
            question_periods.extend([int(m) for m in matches])

        if question_periods:
            existing_periods = []
            for sheet in sheets_data:
                for col in sheet.get("columns", []):
                    col_lower = str(col).lower()
                    for pattern in period_patterns:
                        matches = re.findall(pattern, col_lower)
                        existing_periods.extend([int(m) for m in matches])

            if existing_periods:
                max_existing = max(existing_periods)
                max_asked = max(question_periods)
                if max_asked > max_existing:
                    is_forecast = True

    print(f"IS_FORECAST: {is_forecast}")
    print(f"QUESTION: {question_lower}")

    if is_forecast:
        try:
            forecast_data = run_forecast(sheets_data, periods=2)
            print(f"FORECAST RESULT: {forecast_data}")
            if forecast_data:
                evidence["forecasts"] = forecast_data
        except Exception:
            pass

    print(f"FORECASTS IN EVIDENCE: {'forecasts' in evidence}")
    system_prompt = SYSTEM_PROMPT
    if connections_context:
        system_prompt += (
            "\n\nIMPORTANT: The evidence contains a "
            "'human_confirmed_context' field with "
            "verified relationships between sheets "
            "and columns confirmed by a human analyst. "
            "Always use this context when answering. "
            "When multiple sheets contain the same "
            "metric, prefer the sheet with the highest "
            "confidence score as the primary source."
        )
    if is_forecast:
        system_prompt += """

Forecast projections have been calculated
using linear regression on historical values.
R-squared indicates trend strength (1.0 = perfect).
Present projections as estimates with confidence.
You may also reason about what the trend implies
for business performance beyond just the numbers."""

    # STEP 6 - Send to Gemini
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key or not api_key.strip():
        raise HTTPException(500, detail="GEMINI_API_KEY not configured")

    client = genai.Client(api_key=api_key.strip())

    # Build conversation context
    conversation_context = ""
    if conversation_history:
        conversation_context = "\n\nCONVERSATION HISTORY:\n"
        for turn in conversation_history:
            role = "User" if turn["role"] == "user" else "Assistant"
            conversation_context += f"{role}: {turn['content']}\n"
        conversation_context += "\n"

    user_message = f"""Data from spreadsheet:

{json.dumps(evidence, indent=2)}
{conversation_context}
Current question: {question}"""

    print("EVIDENCE SENT TO GEMINI:")
    print(json.dumps(evidence, indent=2))

    response = client.models.generate_content(
        model="gemini-2.5-flash-lite",
        contents=user_message,
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            max_output_tokens=1024,
        ),
    )
    answer = response.text or "I could not generate a response."
    sheets_used = [s["sheet_name"] for s in selected]

    return answer.strip(), sheets_used


@router.post("/query", response_model=ChatQueryResponse)
async def chat_query(request: ChatQueryRequest):
    """
    Query workbook data via Gemini. Loads metadata, schema mappings, quality issues,
    selects top 3 relevant sheets, builds evidence, and returns answer.
    Supports conversation memory for follow-up questions.
    """
    try:
        workbook_id = request.workbook_id
        question = request.question.strip()
        if not question:
            raise HTTPException(400, detail="question is required")

        # Get or create conversation ID
        if not request.conversation_id:
            conversation_id = secrets.token_hex(8)
        else:
            conversation_id = request.conversation_id

        # Load conversation history
        history = get_conversation_history(conversation_id, _db_path())

        workbook = get_workbook(workbook_id)
        if workbook is None:
            raise HTTPException(404, detail=f"Workbook {workbook_id} not found")

        file_name = workbook.get("file_name", "Unknown")
        sheets = workbook.get("sheets", [])
        schema_mappings = get_schema_mappings(workbook_id)

        quality_report = get_quality_report(workbook_id)
        if quality_report is None:
            quality_report = run_quality_checks(
                workbook_id, sheets, schema_mappings
            )
        quality_issues = quality_report.get("issues", [])

        answer, sheets_used = _build_evidence_and_send(
            workbook_id, file_name, sheets, schema_mappings, quality_issues,
            question, conversation_history=history,
        )

        # Save the conversation turn
        save_conversation_turn(
            conversation_id, workbook_id, question, answer, _db_path()
        )

        return ChatQueryResponse(
            answer=answer,
            sheets_used=sheets_used,
            workbook_id=workbook_id,
            conversation_id=conversation_id,
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, detail=str(e))


@router.get("/conversations/{conversation_id}")
def get_conversation(conversation_id: str):
    """Get conversation history (useful for debugging)."""
    history = get_conversation_history(conversation_id, _db_path(), max_turns=50)
    return {
        "conversation_id": conversation_id,
        "turns": len(history),
        "history": history,
    }


@router.delete("/conversations/{conversation_id}")
def clear_conversation(conversation_id: str):
    """Clear a conversation's history."""
    try:
        conn = sqlite3.connect(_db_path())
        cursor = conn.cursor()
        cursor.execute("""
            DELETE FROM conversations
            WHERE conversation_id = ?
        """, (conversation_id,))
        conn.commit()
        conn.close()
        return {"status": "cleared", "conversation_id": conversation_id}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
