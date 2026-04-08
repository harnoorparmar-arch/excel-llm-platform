import json
import math
import sqlite3
import uuid
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import pandas as pd

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import JSONResponse, Response
from pydantic import BaseModel

from parser.hitl_engine import detect_review_items
from parser.quality_engine import run_quality_checks
from parser.schema_mapper import map_schema
from parser.workbook_parser import (
    _sanitize_for_json,
    parse_workbook,
    SUPPORTED_EXTENSIONS,
)
from storage.database import (
    get_quality_report,
    get_schema_mappings,
    get_workbook,
    init_db,
    store_quality_report,
    store_schema_mapping,
    store_workbook,
)

router = APIRouter(prefix="/workbooks", tags=["workbooks"])


def _db_path() -> str:
    import os
    from pathlib import Path
    path = Path(os.getenv("DATABASE_PATH", "./data/platform.db"))
    path.parent.mkdir(parents=True, exist_ok=True)
    return str(path)


class Decision(BaseModel):
    review_id: str
    type: str
    decision: str
    notes: Optional[str] = None


class DecisionsRequest(BaseModel):
    decisions: List[Decision]

ROOT = Path(__file__).resolve().parent.parent.parent  # excel-llm-platform/
UPLOADS_DIR = ROOT / "uploads"

ALLOWED_EXTENSIONS = ".xlsx .xlsm .xls .xlsb .csv .tsv .ods"
UNSUPPORTED_ERROR = "Unsupported file type. Accepted: .xlsx .xlsm .xls .xlsb .csv .tsv .ods"


def sanitize_for_json(obj):
    """
    Recursively replace NaN, Infinity and
    other non-JSON-serializable floats
    with None so FastAPI can serialize.
    """
    if isinstance(obj, dict):
        return {
            k: sanitize_for_json(v)
            for k, v in obj.items()
        }
    elif isinstance(obj, list):
        return [sanitize_for_json(v) for v in obj]
    elif isinstance(obj, float):
        if math.isnan(obj) or math.isinf(obj):
            return None
        return obj
    else:
        # Handle numpy scalars (e.g. np.float64('nan'))
        try:
            if hasattr(obj, "item"):
                v = obj.item()
                if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
                    return None
            if pd.isna(obj):
                return None
        except (TypeError, ValueError, AttributeError):
            pass
        return obj


@router.post("/upload")
async def upload_workbook(file: UploadFile = File(...)):
    """
    Accept spreadsheet file upload, save to uploads/, parse, store in SQLite,
    return workbook_id and parsed structure.
    """
    if not file.filename:
        raise HTTPException(
            400,
            detail=f"Unsupported file type. Accepted: {ALLOWED_EXTENSIONS}",
        )
    ext = Path(file.filename).suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise HTTPException(
            400,
            detail=f"Unsupported file type. Accepted: {ALLOWED_EXTENSIONS}",
        )

    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    safe_name = f"{uuid.uuid4().hex}_{file.filename}"
    file_path = UPLOADS_DIR / safe_name

    content = await file.read()
    file_path.write_bytes(content)

    parsed = parse_workbook(file_path)
    workbook_id = parsed["workbook_id"]
    file_name = parsed["file_name"]

    store_workbook(
        workbook_id=workbook_id,
        file_name=file_name,
        file_path=str(file_path),
        parsed=parsed,
    )

    payload = {
        "workbook_id": workbook_id,
        "parsed": parsed,
    }
    clean_payload = sanitize_for_json(payload)
    try:
        body = json.dumps(clean_payload)
    except ValueError as e:
        raise HTTPException(500, detail=f"JSON serialization failed: {e}")
    return Response(content=body, media_type="application/json")


@router.post("/{workbook_id}/map-schemas")
async def map_workbook_schemas(workbook_id: str):
    """Load workbook, run map_schema on each non-empty sheet, store in SQLite, return all mappings."""
    workbook = get_workbook(workbook_id)
    if workbook is None:
        raise HTTPException(404, detail=f"Workbook {workbook_id} not found")

    mappings = []
    for sheet in workbook.get("sheets", []):
        columns = sheet.get("columns") or []
        sample_rows = sheet.get("sample_rows") or []
        if not columns or not sample_rows:
            continue
        result = map_schema(
            sheet_name=sheet["sheet_name"],
            columns=columns,
            sample_rows=sample_rows[:10],
            workbook_id=workbook_id,
        )
        store_schema_mapping(workbook_id, sheet["sheet_name"], result)
        mappings.append({**result, "sheet_name": sheet["sheet_name"]})

    return mappings


@router.get("/{workbook_id}/schema-mappings")
async def get_workbook_schema_mappings(workbook_id: str):
    """Return stored schema mappings for a workbook."""
    workbook = get_workbook(workbook_id)
    if workbook is None:
        raise HTTPException(404, detail=f"Workbook {workbook_id} not found")
    return get_schema_mappings(workbook_id)


@router.get("/{workbook_id}/quality-report")
async def get_workbook_quality_report(workbook_id: str):
    """Run quality checks, store in SQLite, return the quality report (overall_score as float)."""
    workbook = get_workbook(workbook_id)
    if workbook is None:
        raise HTTPException(404, detail=f"Workbook {workbook_id} not found")
    sheets_data = workbook.get("sheets", [])
    schema_mappings = get_schema_mappings(workbook_id)
    report = run_quality_checks(workbook_id, sheets_data, schema_mappings)
    store_quality_report(workbook_id, report)
    return {
        **report,
        "overall_score": float(report["overall_score"]),
    }


@router.get("/{workbook_id}")
async def get_workbook_metadata(workbook_id: str):
    """Return stored workbook metadata and parsed structure from SQLite."""
    result = get_workbook(workbook_id)
    if result is None:
        raise HTTPException(404, detail=f"Workbook {workbook_id} not found")
    return result


@router.get("/{workbook_id}/review-items")
def get_review_items(workbook_id: str):
    try:
        init_db()
        db_path = _db_path()
        items = detect_review_items(workbook_id, db_path)
        return {
            "workbook_id": workbook_id,
            "total_items": len(items),
            "high_priority": len([i for i in items if i["priority"] == "high"]),
            "medium_priority": len([i for i in items if i["priority"] == "medium"]),
            "low_priority": len([i for i in items if i["priority"] == "low"]),
            "items": items,
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/{workbook_id}/review-decisions")
def save_review_decisions(workbook_id: str, body: DecisionsRequest):
    try:
        init_db()
        db_path = _db_path()
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        connections_created = 0

        for d in body.decisions:
            cursor.execute(
                """
                INSERT INTO review_decisions
                (workbook_id, review_id, type, decision, notes, decided_at)
                VALUES (?, ?, ?, ?, ?, ?)
            """,
                (
                    workbook_id,
                    d.review_id,
                    d.type,
                    d.decision,
                    d.notes,
                    datetime.now().isoformat(),
                ),
            )

            if d.type == "column_identity" and d.decision.lower().startswith("yes"):
                try:
                    if d.notes:
                        details = json.loads(d.notes)
                        cursor.execute(
                            """
                            INSERT INTO sheet_connections
                            (workbook_id, sheet_1, identifier_1, sheet_2,
                             identifier_2, connection_type, confirmed, confirmed_at)
                            VALUES (?,?,?,?,?,?,1,?)
                        """,
                            (
                                workbook_id,
                                details.get("sheet_1"),
                                details.get("column_1"),
                                details.get("sheet_2"),
                                details.get("column_2"),
                                "column_match",
                                datetime.now().isoformat(),
                            ),
                        )
                        connections_created += 1
                except Exception:
                    pass

            if (
                d.type == "same_metric_across_sheets"
                and d.decision.lower().startswith("yes")
            ):
                try:
                    if d.notes:
                        details = json.loads(d.notes)
                        appearances = details.get("appearances", [])
                        metric = details.get("metric_name", "")

                        for i in range(len(appearances)):
                            for j in range(i + 1, len(appearances)):
                                cursor.execute(
                                    """
                                    INSERT INTO sheet_connections
                                    (workbook_id, sheet_1, identifier_1, sheet_2,
                                     identifier_2, connection_type, confirmed, confirmed_at)
                                    VALUES (?,?,?,?,?,?,1,?)
                                """,
                                    (
                                        workbook_id,
                                        appearances[i]["sheet"],
                                        metric,
                                        appearances[j]["sheet"],
                                        metric,
                                        "same_metric",
                                        datetime.now().isoformat(),
                                    ),
                                )
                                connections_created += 1
                except Exception:
                    pass

        conn.commit()
        conn.close()

        return {
            "status": "decisions_saved",
            "decisions_count": len(body.decisions),
            "connections_created": connections_created,
            "workbook_id": workbook_id,
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/{workbook_id}/connections")
def get_connections(workbook_id: str):
    try:
        init_db()
        db_path = _db_path()
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute("""
            SELECT sheet_1, identifier_1,
                   sheet_2, identifier_2,
                   connection_type,
                   confirmed_at
            FROM sheet_connections
            WHERE workbook_id = ?
            AND confirmed = 1
            ORDER BY confirmed_at
        """, (workbook_id,))

        connections = [dict(row) for row in cursor.fetchall()]
        conn.close()

        return {
            "workbook_id": workbook_id,
            "connections": connections,
            "total": len(connections),
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/{workbook_id}/review-status")
def get_review_status(workbook_id: str):
    try:
        init_db()
        db_path = _db_path()
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT COUNT(*) FROM review_decisions
            WHERE workbook_id = ?
        """,
            (workbook_id,),
        )
        decided = cursor.fetchone()[0]

        cursor.execute(
            """
            SELECT COUNT(*) FROM sheet_connections
            WHERE workbook_id = ? AND confirmed = 1
        """,
            (workbook_id,),
        )
        connections = cursor.fetchone()[0]

        conn.close()

        items = detect_review_items(workbook_id, db_path)
        total = len(items)
        high_priority = len([i for i in items if i["priority"] == "high"])
        high_decided = min(decided, high_priority)
        ready = high_decided >= high_priority

        return {
            "workbook_id": workbook_id,
            "total_items": total,
            "decided": decided,
            "pending": max(0, total - decided),
            "connections_confirmed": connections,
            "ready_for_analysis": ready,
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
