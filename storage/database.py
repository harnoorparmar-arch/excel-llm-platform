"""SQLite-backed storage for workbooks and chat state."""
import json
import os
import sqlite3
from pathlib import Path
from typing import Any


def _db_path() -> Path:
    path = Path(os.getenv("DATABASE_PATH", "./data/platform.db"))
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def _get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(str(_db_path()))
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    """Create workbooks and schema_mappings tables if they do not exist."""
    with _get_connection() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS workbooks (
                workbook_id TEXT PRIMARY KEY,
                file_name TEXT NOT NULL,
                file_path TEXT NOT NULL,
                parsed_json TEXT NOT NULL,
                created_at TEXT DEFAULT (datetime('now'))
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS schema_mappings (
                workbook_id TEXT NOT NULL,
                sheet_name TEXT NOT NULL,
                entity TEXT NOT NULL,
                time_structure TEXT,
                mapped_fields TEXT,
                key_metrics TEXT,
                confidence REAL NOT NULL,
                needs_review INTEGER NOT NULL,
                error TEXT,
                created_at TEXT DEFAULT (datetime('now')),
                PRIMARY KEY (workbook_id, sheet_name)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS quality_reports (
                workbook_id TEXT PRIMARY KEY,
                overall_score REAL NOT NULL,
                total_issues INTEGER NOT NULL,
                issues_json TEXT NOT NULL,
                created_at TEXT DEFAULT (datetime('now'))
            )
        """)
        # Migration: fix existing tables where overall_score was INTEGER (truncates 0.16 -> 0)
        try:
            for row in conn.execute("PRAGMA table_info(quality_reports)").fetchall():
                if row[1] == "overall_score" and str(row[2]).upper() != "REAL":
                    conn.execute("DROP TABLE quality_reports")
                    conn.execute("""
                        CREATE TABLE quality_reports (
                            workbook_id TEXT PRIMARY KEY,
                            overall_score REAL NOT NULL,
                            total_issues INTEGER NOT NULL,
                            issues_json TEXT NOT NULL,
                            created_at TEXT DEFAULT (datetime('now'))
                        )
                    """)
                    break
        except Exception:
            pass
        conn.execute("""
            CREATE TABLE IF NOT EXISTS review_decisions (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                workbook_id     TEXT NOT NULL,
                review_id       TEXT NOT NULL,
                type            TEXT NOT NULL,
                decision        TEXT NOT NULL,
                notes           TEXT,
                decided_by      TEXT DEFAULT 'human',
                decided_at      TEXT NOT NULL,
                auto_confirmed  INTEGER DEFAULT 0
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS sheet_connections (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                workbook_id     TEXT NOT NULL,
                sheet_1         TEXT NOT NULL,
                identifier_1    TEXT NOT NULL,
                sheet_2         TEXT NOT NULL,
                identifier_2    TEXT NOT NULL,
                connection_type TEXT NOT NULL,
                confirmed       INTEGER DEFAULT 0,
                confirmed_by    TEXT DEFAULT 'human',
                confirmed_at    TEXT,
                notes           TEXT
            )
        """)
        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_review_decisions_workbook
            ON review_decisions(workbook_id)
        """)
        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_sheet_connections_workbook
            ON sheet_connections(workbook_id)
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS workspaces (
                workspace_id  TEXT PRIMARY KEY,
                name          TEXT,
                created_at    TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS workspace_files (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                workspace_id  TEXT NOT NULL,
                workbook_id   TEXT NOT NULL,
                file_name     TEXT NOT NULL,
                added_at      TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS cross_file_relationships (
                id                INTEGER PRIMARY KEY AUTOINCREMENT,
                workspace_id      TEXT NOT NULL,
                file_1            TEXT NOT NULL,
                column_1          TEXT NOT NULL,
                file_2            TEXT NOT NULL,
                column_2          TEXT NOT NULL,
                confidence        REAL,
                relationship_type TEXT,
                is_primary_key    INTEGER DEFAULT 0,
                confirmed         INTEGER DEFAULT 0,
                confirmed_by      TEXT,
                confirmed_at      TEXT,
                rejected         INTEGER DEFAULT 0
            )
        """)
        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_workspace_files_workspace
            ON workspace_files(workspace_id)
        """)
        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_cross_file_relationships_workspace
            ON cross_file_relationships(workspace_id)
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS conversations (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                conversation_id TEXT NOT NULL,
                workbook_id     TEXT NOT NULL,
                role            TEXT NOT NULL,
                content         TEXT NOT NULL,
                created_at      TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_conversations_conversation_id
            ON conversations(conversation_id)
        """)
        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_conversations_workbook_id
            ON conversations(workbook_id)
        """)


def store_workbook(
    workbook_id: str,
    file_name: str,
    file_path: str,
    parsed: dict[str, Any],
) -> None:
    """Store workbook metadata and parsed structure in SQLite."""
    init_db()
    with _get_connection() as conn:
        conn.execute(
            """
            INSERT INTO workbooks (workbook_id, file_name, file_path, parsed_json)
            VALUES (?, ?, ?, ?)
            """,
            (workbook_id, file_name, file_path, json.dumps(parsed)),
        )


def get_workbook(workbook_id: str) -> dict[str, Any] | None:
    """Return stored workbook metadata and parsed structure, or None if not found."""
    init_db()
    with _get_connection() as conn:
        row = conn.execute(
            "SELECT workbook_id, file_name, file_path, parsed_json, created_at FROM workbooks WHERE workbook_id = ?",
            (workbook_id,),
        ).fetchone()
    if row is None:
        return None
    parsed = json.loads(row["parsed_json"])
    return {
        "workbook_id": row["workbook_id"],
        "file_name": row["file_name"],
        "file_path": row["file_path"],
        "created_at": row["created_at"],
        **parsed,
    }


def store_schema_mapping(
    workbook_id: str,
    sheet_name: str,
    mapping: dict[str, Any],
) -> None:
    """Store or replace a schema mapping for a workbook sheet."""
    init_db()
    entity = mapping.get("entity", "unknown")
    time_structure = mapping.get("time_structure") or ""
    mapped_fields = json.dumps(mapping.get("mapped_fields") or {})
    key_metrics = json.dumps(mapping.get("key_metrics") or [])
    confidence = float(mapping.get("confidence", 0))
    needs_review = 1 if mapping.get("needs_review") else 0
    error = mapping.get("error")

    with _get_connection() as conn:
        conn.execute(
            """
            INSERT OR REPLACE INTO schema_mappings
            (workbook_id, sheet_name, entity, time_structure, mapped_fields, key_metrics, confidence, needs_review, error)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (workbook_id, sheet_name, entity, time_structure, mapped_fields, key_metrics, confidence, needs_review, error),
        )


def get_schema_mappings(workbook_id: str) -> list[dict[str, Any]]:
    """Return all schema mappings for a workbook."""
    init_db()
    with _get_connection() as conn:
        rows = conn.execute(
            """
            SELECT workbook_id, sheet_name, entity, time_structure, mapped_fields, key_metrics, confidence, needs_review, error
            FROM schema_mappings WHERE workbook_id = ? ORDER BY sheet_name
            """,
            (workbook_id,),
        ).fetchall()
    return [
        {
            "workbook_id": row["workbook_id"],
            "sheet_name": row["sheet_name"],
            "entity": row["entity"],
            "time_structure": row["time_structure"] or None,
            "mapped_fields": json.loads(row["mapped_fields"] or "{}"),
            "key_metrics": json.loads(row["key_metrics"] or "[]"),
            "confidence": row["confidence"],
            "needs_review": bool(row["needs_review"]),
            "error": row["error"],
        }
        for row in rows
    ]


def store_quality_report(workbook_id: str, report: dict[str, Any]) -> None:
    """Store or replace a quality report. overall_score stored as REAL (float)."""
    init_db()
    overall_score = float(report.get("overall_score", 0.0))
    total_issues = int(report.get("total_issues", 0))
    issues_json = json.dumps(report.get("issues", []))
    with _get_connection() as conn:
        conn.execute(
            """
            INSERT OR REPLACE INTO quality_reports
            (workbook_id, overall_score, total_issues, issues_json)
            VALUES (?, ?, ?, ?)
            """,
            (workbook_id, overall_score, total_issues, issues_json),
        )


def get_quality_report(workbook_id: str) -> dict[str, Any] | None:
    """Return stored quality report, or None. overall_score as float."""
    init_db()
    with _get_connection() as conn:
        row = conn.execute(
            "SELECT workbook_id, overall_score, total_issues, issues_json, created_at FROM quality_reports WHERE workbook_id = ?",
            (workbook_id,),
        ).fetchone()
    if row is None:
        return None
    return {
        "workbook_id": row["workbook_id"],
        "overall_score": float(row["overall_score"]),
        "total_issues": row["total_issues"],
        "issues": json.loads(row["issues_json"]),
        "created_at": row["created_at"],
    }
