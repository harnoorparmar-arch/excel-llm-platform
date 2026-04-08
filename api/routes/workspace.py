"""Workspace routes - groups of related files analyzed together."""
import os
import json
import re
import sqlite3
import secrets
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from fastapi import APIRouter, HTTPException, UploadFile, File
from pydantic import BaseModel
from typing import Optional

from storage.database import init_db, get_workbook, get_schema_mappings, store_workbook, store_quality_report
from parser.workbook_parser import parse_workbook, SUPPORTED_EXTENSIONS
from parser.relationship_engine import (
    detect_cross_file_relationships,
    save_relationships_to_db,
)
from parser.schema_mapper import map_schema
from parser.quality_engine import run_quality_checks
from parser.unification_engine import run_unification

router = APIRouter(prefix="/workspaces", tags=["workspaces"])


def _db_path() -> str:
    path = Path(os.getenv("DATABASE_PATH", "./data/platform.db"))
    path.parent.mkdir(parents=True, exist_ok=True)
    return str(path)


def _upload_dir() -> str:
    root = Path(__file__).resolve().parent.parent.parent
    upload_dir = root / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)
    return str(upload_dir)


def _init_workspace_tables(conn: sqlite3.Connection) -> None:
    """Create workspaces and workspace_files tables if they do not exist."""
    conn.execute("""
        CREATE TABLE IF NOT EXISTS workspaces (
            workspace_id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS workspace_files (
            workspace_id TEXT NOT NULL,
            workbook_id TEXT NOT NULL,
            file_name TEXT NOT NULL,
            added_at TEXT NOT NULL,
            PRIMARY KEY (workspace_id, workbook_id)
        )
    """)
    conn.commit()


# ---------------------------------------------------------------------------
# ENDPOINT 1: POST /workspaces/create
# ---------------------------------------------------------------------------
@router.post("/create")
def create_workspace(name: Optional[str] = None):
    init_db()
    workspace_id = secrets.token_hex(8)
    workspace_name = name or f"Workspace {workspace_id[:6]}"

    conn = sqlite3.connect(_db_path())
    _init_workspace_tables(conn)
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO workspaces
        (workspace_id, name, created_at)
        VALUES (?, ?, ?)
    """, (
        workspace_id,
        workspace_name,
        datetime.now().isoformat()
    ))

    conn.commit()
    conn.close()

    return {
        "workspace_id": workspace_id,
        "name": workspace_name,
        "created_at": datetime.now().isoformat()
    }


# ---------------------------------------------------------------------------
# ENDPOINT 2: POST /workspaces/{workspace_id}/upload
# ---------------------------------------------------------------------------
@router.post("/{workspace_id}/upload")
async def upload_to_workspace(
    workspace_id: str,
    file: UploadFile = File(...)
):
    init_db()

    # Verify workspace exists
    conn = sqlite3.connect(_db_path())
    _init_workspace_tables(conn)
    cursor = conn.cursor()
    cursor.execute(
        "SELECT workspace_id FROM workspaces WHERE workspace_id = ?",
        (workspace_id,)
    )
    if not cursor.fetchone():
        conn.close()
        raise HTTPException(
            status_code=404,
            detail="Workspace not found"
        )
    conn.close()

    if not file.filename:
        raise HTTPException(400, detail="No filename")

    ext = Path(file.filename).suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise HTTPException(
            400,
            detail=f"Unsupported file type. Accepted: .xlsx .xlsm .xls .xlsb .csv .tsv .ods"
        )

    # Save file to disk (same pattern as upload.py)
    safe_name = f"{secrets.token_hex(8)}_{file.filename}"
    file_path = os.path.join(_upload_dir(), safe_name)

    contents = await file.read()
    with open(file_path, "wb") as f:
        f.write(contents)

    # Parse using existing parser
    parsed = parse_workbook(file_path)
    workbook_id = parsed["workbook_id"]
    file_name = parsed["file_name"]

    # Store in workbooks table (reuse storage.database)
    store_workbook(
        workbook_id=workbook_id,
        file_name=file_name,
        file_path=file_path,
        parsed=parsed,
    )

    # Link file to workspace
    conn = sqlite3.connect(_db_path())
    _init_workspace_tables(conn)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT OR REPLACE INTO workspace_files
        (workspace_id, workbook_id, file_name, added_at)
        VALUES (?, ?, ?, ?)
    """, (
        workspace_id,
        workbook_id,
        file.filename,
        datetime.now().isoformat()
    ))
    conn.commit()
    conn.close()

    # Run schema mapping (reuse logic from map_workbook_schemas)
    workbook = get_workbook(workbook_id)
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
        from storage.database import store_schema_mapping
        store_schema_mapping(workbook_id, sheet["sheet_name"], result)

    # Run quality check
    sheets_data = workbook.get("sheets", [])
    schema_mappings = get_schema_mappings(workbook_id)
    report = run_quality_checks(workbook_id, sheets_data, schema_mappings)
    store_quality_report(workbook_id, report)

    sheet_count = len(parsed.get("sheets", []))

    return {
        "workbook_id": workbook_id,
        "workspace_id": workspace_id,
        "file_name": file.filename,
        "sheet_count": sheet_count,
        "status": "processed"
    }


# ---------------------------------------------------------------------------
# ENDPOINT 3: GET /workspaces/{workspace_id}/files
# ---------------------------------------------------------------------------
@router.get("/{workspace_id}/files")
def get_workspace_files(workspace_id: str):
    conn = sqlite3.connect(_db_path())
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    _init_workspace_tables(conn)

    # Get workspace
    cursor.execute(
        "SELECT * FROM workspaces WHERE workspace_id = ?",
        (workspace_id,)
    )
    workspace = cursor.fetchone()
    if not workspace:
        conn.close()
        raise HTTPException(
            status_code=404,
            detail="Workspace not found"
        )

    # Get all files (workbooks and workspace_files in same db)
    cursor.execute("""
        SELECT wf.workbook_id, wf.file_name,
               wf.added_at,
               w.parsed_json
        FROM workspace_files wf
        JOIN workbooks w
          ON wf.workbook_id = w.workbook_id
        WHERE wf.workspace_id = ?
        ORDER BY wf.added_at
    """, (workspace_id,))
    files = cursor.fetchall()

    result = []
    for f in files:
        # Get schema mappings
        cursor.execute("""
            SELECT sheet_name, entity,
                   confidence, key_metrics
            FROM schema_mappings
            WHERE workbook_id = ?
        """, (f["workbook_id"],))
        mappings = [dict(m) for m in cursor.fetchall()]

        # Get quality score
        cursor.execute("""
            SELECT overall_score, total_issues
            FROM quality_reports
            WHERE workbook_id = ?
        """, (f["workbook_id"],))
        quality = cursor.fetchone()

        parsed = json.loads(f["parsed_json"] or "{}")

        result.append({
            "workbook_id": f["workbook_id"],
            "file_name": f["file_name"],
            "added_at": f["added_at"],
            "sheet_count": len(parsed.get("sheets", [])),
            "schema_mappings": mappings,
            "quality_score": float(quality["overall_score"]) if quality else None,
            "total_issues": quality["total_issues"] if quality else None
        })

    conn.close()

    return {
        "workspace_id": workspace_id,
        "name": dict(workspace)["name"],
        "total_files": len(result),
        "files": result
    }


# ---------------------------------------------------------------------------
# ENDPOINT 4: GET /workspaces
# ---------------------------------------------------------------------------
@router.get("")
def list_workspaces():
    conn = sqlite3.connect(_db_path())
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    _init_workspace_tables(conn)

    cursor.execute("""
        SELECT w.workspace_id, w.name,
               w.created_at,
               COUNT(wf.workbook_id) as file_count
        FROM workspaces w
        LEFT JOIN workspace_files wf
          ON w.workspace_id = wf.workspace_id
        GROUP BY w.workspace_id
        ORDER BY w.created_at DESC
    """)

    workspaces = [dict(row) for row in cursor.fetchall()]
    conn.close()

    return {
        "total": len(workspaces),
        "workspaces": workspaces
    }


# ---------------------------------------------------------------------------
# ENDPOINT 5: POST /workspaces/{workspace_id}/detect-relationships
# ---------------------------------------------------------------------------
@router.post("/{workspace_id}/detect-relationships")
def detect_relationships(workspace_id: str):
    # Verify workspace exists
    conn = sqlite3.connect(_db_path())
    cursor = conn.cursor()
    cursor.execute(
        "SELECT workspace_id FROM workspaces WHERE workspace_id = ?",
        (workspace_id,)
    )
    if not cursor.fetchone():
        conn.close()
        raise HTTPException(
            status_code=404,
            detail="Workspace not found"
        )
    conn.close()

    # Run detection
    result = detect_cross_file_relationships(
        workspace_id, _db_path()
    )

    # Save to database
    if result.get('relationships'):
        save_relationships_to_db(
            workspace_id,
            result['relationships'],
            _db_path()
        )

    return result


# ---------------------------------------------------------------------------
# ENDPOINT 6: GET /workspaces/{workspace_id}/relationships
# ---------------------------------------------------------------------------
@router.get("/{workspace_id}/relationships")
def get_relationships(workspace_id: str):
    conn = sqlite3.connect(_db_path())
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        SELECT * FROM cross_file_relationships
        WHERE workspace_id = ?
        ORDER BY confidence DESC
    """, (workspace_id,))

    relationships = [
        dict(row) for row in cursor.fetchall()
    ]
    conn.close()

    return {
        "workspace_id": workspace_id,
        "total": len(relationships),
        "relationships": relationships
    }


# ---------------------------------------------------------------------------
# ENDPOINT 7: POST /workspaces/{workspace_id}/relationships/{rel_id}/confirm
# ---------------------------------------------------------------------------
@router.post("/{workspace_id}/relationships/{rel_id}/confirm")
def confirm_relationship(
    workspace_id: str,
    rel_id: int
):
    conn = sqlite3.connect(_db_path())
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE cross_file_relationships
        SET confirmed = 1,
            confirmed_by = 'human',
            confirmed_at = ?
        WHERE id = ? AND workspace_id = ?
    """, (
        datetime.now().isoformat(),
        rel_id,
        workspace_id
    ))
    conn.commit()
    conn.close()
    return {"status": "confirmed", "id": rel_id}


# ---------------------------------------------------------------------------
# ENDPOINT 8: POST /workspaces/{workspace_id}/relationships/{rel_id}/reject
# ---------------------------------------------------------------------------
@router.post("/{workspace_id}/relationships/{rel_id}/reject")
def reject_relationship(
    workspace_id: str,
    rel_id: int
):
    conn = sqlite3.connect(_db_path())
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE cross_file_relationships
        SET rejected = 1
        WHERE id = ? AND workspace_id = ?
    """, (rel_id, workspace_id))
    conn.commit()
    conn.close()
    return {"status": "rejected", "id": rel_id}


# ---------------------------------------------------------------------------
# ENDPOINT 9: POST /workspaces/{workspace_id}/unify
# ---------------------------------------------------------------------------
@router.post("/{workspace_id}/unify")
def unify_workspace(workspace_id: str):
    # Verify workspace exists
    conn = sqlite3.connect(_db_path())
    cursor = conn.cursor()
    cursor.execute(
        "SELECT workspace_id FROM workspaces WHERE workspace_id = ?",
        (workspace_id,)
    )
    if not cursor.fetchone():
        conn.close()
        raise HTTPException(
            status_code=404,
            detail="Workspace not found"
        )
    conn.close()

    # Run unification
    report = run_unification(
        workspace_id, _db_path()
    )

    return report


# ---------------------------------------------------------------------------
# ENDPOINT 10: GET /workspaces/{workspace_id}/unified-data
# ---------------------------------------------------------------------------
@router.get("/{workspace_id}/unified-data")
def get_workspace_unified_data(workspace_id: str):
    conn = sqlite3.connect(_db_path())
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        SELECT file_name, unified_data, created_at
        FROM unified_datasets
        WHERE workspace_id = ?
        ORDER BY file_name
    """, (workspace_id,))

    results = []
    for row in cursor.fetchall():
        data = json.loads(row['unified_data'])
        results.append({
            'file_name': row['file_name'],
            'created_at': row['created_at'],
            'clean_rows': data.get('clean_count', 0),
            'columns': data.get('columns', []),
            'sample_rows': data.get('rows', [])[:5]
        })

    conn.close()

    return {
        'workspace_id': workspace_id,
        'files': results,
        'total_files': len(results)
    }


@router.get("/{workspace_id}/debug-aggs")
def debug_aggregations(workspace_id: str):
    conn = sqlite3.connect(_db_path())
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("""
        SELECT file_name, unified_data
        FROM unified_datasets WHERE workspace_id = ?
    """, (workspace_id,))
    unified_files = []
    for row in cursor.fetchall():
        data = json.loads(row['unified_data'])
        unified_files.append({
            'file': row['file_name'],
            'columns': data.get('columns', []),
            'rows': data.get('rows', [])[:200]
        })
    conn.close()
    aggs = compute_aggregations(unified_files)
    result = {}
    for file, data in aggs.items():
        result[file] = {
            'column_types': data.get('_column_types', {}),
            'group_performance_keys': list(
                data.get('group_performance', {}).keys()),
            'time_series_keys': list(
                data.get('time_series', {}).keys()),
            'rankings_index': data.get('_rankings_index', {}),
            'timeseries_index': data.get('_timeseries_index', {}),
            'top_group_per_key': {
                k: {
                    'top': v.get('top'),
                    'top_total': v.get('top_total')
                }
                for k, v in data.get(
                    'group_performance', {}).items()
            },
            'time_series_best_months': {
                k: {
                    'best_month': v.get('best_month'),
                    'best_month_value': v.get(
                        'best_month_value'),
                    'note': v.get('note'),
                    'date_col': v.get('date_column_used')
                }
                for k, v in data.get(
                    'time_series', {}).items()
            }
        }
    return result


@router.get("/{workspace_id}/debug-unified")
def debug_unified(workspace_id: str):
    conn = sqlite3.connect(_db_path())
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("""
        SELECT file_name, unified_data
        FROM unified_datasets
        WHERE workspace_id = ?
        ORDER BY file_name
    """, (workspace_id,))

    result = {}
    for row in cursor.fetchall():
        data = json.loads(row['unified_data'])
        rows = data.get('rows', [])
        columns = data.get('columns', [])

        # For sales_orders specifically check
        # sales_rep totals from stored data
        rep_totals = {}
        if 'sales_rep' in columns or any(
            'sales_rep' in str(c) for c in columns
        ):
            from collections import defaultdict
            totals = defaultdict(float)
            for r in rows:
                rep = str(r.get('sales_rep',
                               r.get('sales_rep', '')
                               )).strip()
                try:
                    price = float(r.get('sale_price', 0) or 0)
                    qty = float(r.get('quantity', 1) or 1)
                    totals[rep] += price * qty
                except Exception:
                    pass
            rep_totals = dict(
                sorted(totals.items(),
                       key=lambda x: -x[1])
            )

        result[row['file_name']] = {
            'total_rows_stored': len(rows),
            'columns': columns,
            'first_row': rows[0] if rows else None,
            'last_row': rows[-1] if rows else None,
            'rep_totals_from_stored_data': rep_totals
        }

    conn.close()
    return result


# ---------------------------------------------------------------------------
# ENDPOINT 11: POST /workspaces/{workspace_id}/chat
# ---------------------------------------------------------------------------
def compute_aggregations(unified_files):
    """
    Pre-computes aggregations from any spreadsheet schema.
    Uses heuristics to detect column types rather than
    hardcoded column names. Produces numeric summaries,
    categorical breakdowns, time series, and group-by
    performance metrics that work across any domain.
    """
    from collections import defaultdict
    from datetime import datetime

    # Heuristics for column type detection
    DATE_HINTS = ['date', 'time', 'day', 'month', 'year',
                  'period', 'when', 'created', 'updated',
                  'ordered', 'invoiced', 'closed', 'due']
    AMOUNT_HINTS = ['amount', 'price', 'cost', 'value',
                    'revenue', 'total', 'sales', 'income',
                    'payment', 'balance', 'fee', 'charge',
                    'salary', 'budget', 'spend', 'profit',
                    'margin', 'earning', 'due', 'subtotal',
                    'balance_due', 'list_price', 'our_price',
                    'unit_cost', 'sale_price', 'margin_pct',
                    'rate', 'wage', 'gross', 'net', 'tax']
    QTY_HINTS = ['quantity', 'qty', 'count', 'units',
                 'volume', 'number', 'num', 'amount']
    GROUP_HINTS = ['rep', 'agent', 'manager', 'owner',
                   'person', 'employee', 'staff', 'user',
                   'category', 'type', 'region', 'area',
                   'segment', 'department', 'team', 'group',
                   'brand', 'vendor', 'supplier', 'source',
                   'name', 'assigned', 'handled', 'territory']
    ID_HINTS = ['id', 'key', 'code', 'number', 'ref',
                'identifier', 'sku', 'no']
    STATUS_HINTS = ['status', 'state', 'stage', 'flag',
                    'type', 'category', 'class']

    def detect_col_type(col_name, sample_values):
        """Classify a column by name hints and sample values."""
        col = str(col_name).lower()

        # Try parsing sample values as dates
        date_parseable = 0
        for v in sample_values[:10]:
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y',
                        '%Y/%m/%d', '%d-%m-%Y']:
                try:
                    datetime.strptime(str(v).strip(), fmt)
                    date_parseable += 1
                    break
                except (ValueError, TypeError):
                    pass

        # Try parsing as numeric
        numeric_count = 0
        for v in sample_values[:10]:
            try:
                float(str(v).replace(',', '').replace('$', ''))
                numeric_count += 1
            except (ValueError, TypeError):
                pass

        n = len(sample_values[:10]) or 1

        if date_parseable / n > 0.5:
            return 'date'
        if any(h in col for h in DATE_HINTS) and date_parseable > 0:
            return 'date'
        if numeric_count / n > 0.7:
            if any(h in col for h in AMOUNT_HINTS):
                return 'amount'
            if any(h in col for h in QTY_HINTS):
                return 'quantity'
            if any(h in col for h in ID_HINTS):
                return 'id'
            return 'numeric'
        if any(h in col for h in STATUS_HINTS):
            return 'status'
        if any(h in col for h in GROUP_HINTS):
            return 'group'
        if any(h in col for h in ID_HINTS):
            return 'id'
        return 'text'

    def safe_float(v):
        try:
            return float(str(v).replace(',', '').replace('$', ''))
        except (ValueError, TypeError):
            return None

    def parse_date(v):
        for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y',
                    '%Y/%m/%d', '%d-%m-%Y']:
            try:
                return datetime.strptime(str(v).strip(), fmt)
            except (ValueError, TypeError):
                pass
        return None

    aggregations = {}

    for file_data in unified_files:
        file_name = file_data['file']
        rows = file_data.get('rows', [])
        columns = file_data.get('columns') or (
            list(rows[0].keys()) if rows else []
        )

        # Separate native columns from columns that
        # were joined in from other files during
        # unification. Joined columns have a file
        # prefix like invoices_, customers_, products_
        # We detect native columns as those that do
        # not start with a known file prefix.
        # This prevents joined columns from polluting
        # column type detection for the native schema.
        known_prefixes = tuple(
            f['file'].replace('.csv', '').replace('.xlsx', '') + '_'
            for f in unified_files
            if f['file'] != file_name
        )
        native_columns = [
            c for c in columns
            if not str(c).startswith(known_prefixes)
        ]
        joined_columns = [
            c for c in columns
            if str(c).startswith(known_prefixes)
        ]

        if not rows:
            continue

        file_aggs = {}

        # --- Classify all columns ---
        col_types = {}
        col_samples = {}
        for col in native_columns:
            samples = [r.get(col) for r in rows[:20]
                      if r.get(col) is not None]
            col_samples[col] = samples
            col_types[col] = detect_col_type(col, samples)
        for col in joined_columns:
            samples = [r.get(col) for r in rows[:20]
                      if r.get(col) is not None]
            col_samples[col] = samples
            col_types[col] = detect_col_type(col, samples)

        date_cols = [c for c in native_columns if col_types.get(c) == 'date']
        amount_cols = [c for c in native_columns if col_types.get(c) == 'amount']
        qty_cols = [c for c in native_columns if col_types.get(c) == 'quantity']
        group_cols = [c for c in native_columns if col_types.get(c) == 'group']
        status_cols = [c for c in native_columns if col_types.get(c) == 'status']
        numeric_cols = [c for c in columns
                       if col_types.get(c) in ('amount', 'quantity', 'numeric')]

        # Store column classification for LLM context
        file_aggs['_column_types'] = {
            'date': date_cols,
            'amount': amount_cols,
            'quantity': qty_cols,
            'group': group_cols,
            'status': status_cols
        }

        # --- 1. Numeric summaries for all numeric columns ---
        numeric_summaries = {}
        for col in numeric_cols:
            values = [safe_float(r.get(col)) for r in rows]
            values = [v for v in values if v is not None]
            if values:
                numeric_summaries[col] = {
                    'sum': round(sum(values), 2),
                    'min': round(min(values), 2),
                    'max': round(max(values), 2),
                    'avg': round(sum(values) / len(values), 2),
                    'count': len(values)
                }
        if numeric_summaries:
            file_aggs['numeric_summaries'] = numeric_summaries

        # --- 2. Categorical breakdowns for group/status columns ---
        categorical = {}
        for col in group_cols + status_cols:
            counts = defaultdict(int)
            amount_by_group = defaultdict(float)
            for row in rows:
                val = str(row.get(col, '') or '').strip()
                if not val:
                    continue
                counts[val] += 1
                # Sum the primary amount column by this group
                if amount_cols:
                    primary_amount = amount_cols[0]
                    amt = safe_float(row.get(primary_amount))
                    if amt is not None:
                        amount_by_group[val] += amt
            if counts:
                sorted_counts = sorted(counts.items(),
                                      key=lambda x: -x[1])
                entry = {
                    'value_counts': dict(sorted_counts[:15]),
                    'unique_count': len(counts),
                    'total_count': sum(counts.values())
                }
                if amount_by_group and amount_cols:
                    sorted_by_amount = sorted(
                        amount_by_group.items(),
                        key=lambda x: -x[1]
                    )
                    entry['total_by_group'] = {
                        k: round(v, 2)
                        for k, v in sorted_by_amount[:15]
                    }
                    entry['top_group'] = sorted_by_amount[0][0]
                    entry['top_group_total'] = round(
                        sorted_by_amount[0][1], 2
                    )
                categorical[col] = entry
        if categorical:
            file_aggs['categorical_breakdowns'] = categorical

        # --- 3. Conditional aggregations: amount by status ---
        if status_cols and amount_cols:
            conditional = {}
            for status_col in status_cols[:2]:
                for amount_col in amount_cols[:3]:
                    by_status = defaultdict(list)
                    for row in rows:
                        sv = str(row.get(status_col, '') or '').strip().lower()
                        av = safe_float(row.get(amount_col))
                        if sv and av is not None:
                            by_status[sv].append(av)
                    if by_status:
                        key = f'{amount_col}_by_{status_col}'
                        conditional[key] = {
                            sv: {
                                'sum': round(sum(vals), 2),
                                'count': len(vals)
                            }
                            for sv, vals in by_status.items()
                        }
            # Build entity-level outstanding balances
            # Only include records where status
            # indicates money is owed
            OWED_STATUSES = {
                'unpaid', 'overdue', 'partial',
                'outstanding', 'pending', 'due',
                'late', 'delinquent'
            }
            for status_col in status_cols[:2]:
                for amount_col in amount_cols[:5]:
                    col_lower = str(amount_col).lower()
                    # Only use balance/due columns
                    # not total or subtotal columns
                    is_balance_col = any(
                        h in col_lower
                        for h in ['balance', 'due',
                                  'outstanding',
                                  'remaining',
                                  'owed']
                    )
                    if not is_balance_col:
                        continue
                    # Find entity name columns
                    name_cols = [
                        c for c in native_columns
                        if col_types.get(c) in
                        ('group',)
                        and any(h in str(c).lower()
                               for h in ['name',
                                        'company',
                                        'customer',
                                        'client',
                                        'vendor'])
                    ]
                    if not name_cols:
                        continue
                    name_col = name_cols[0]
                    entity_balances = defaultdict(
                        float)
                    for row in rows:
                        sv = str(
                            row.get(status_col, '')
                            or ''
                        ).strip().lower()
                        if sv not in OWED_STATUSES:
                            continue
                        entity = str(
                            row.get(name_col, '')
                            or ''
                        ).strip()
                        val = safe_float(
                            row.get(amount_col))
                        if entity and val:
                            entity_balances[
                                entity] += val
                    if entity_balances:
                        ranked = sorted(
                            entity_balances.items(),
                            key=lambda x: -x[1]
                        )
                        key = f'outstanding_{amount_col}_by_{name_col}'
                        conditional[key] = {
                            'total_outstanding': round(
                                sum(entity_balances
                                    .values()), 2),
                            'by_entity': {
                                k: round(v, 2)
                                for k, v in ranked
                            },
                            'top_entity': ranked[0][0],
                            'top_entity_balance':
                                round(ranked[0][1], 2)
                        }
            if conditional:
                file_aggs['conditional_aggregations'] = conditional

        # --- 4. Time series: group any amount column by month ---
        if date_cols and (amount_cols or qty_cols):
            date_col = date_cols[0]
            time_series = {}

            target_cols = amount_cols[:3] + qty_cols[:1]
            for val_col in target_cols:
                col_lower = str(val_col).lower()
                is_per_unit = any(h in col_lower
                    for h in ['unit_price', 'our_price',
                              'list_price', 'sale_price',
                              'unit_cost', 'cost_price'])
                is_already_total = any(h in col_lower
                    for h in ['total', 'subtotal', 'balance',
                              'amount_paid', 'balance_due'])

                monthly = defaultdict(float)
                for row in rows:
                    d = parse_date(row.get(date_col))
                    if not d:
                        continue
                    v = safe_float(row.get(val_col))
                    if v is None:
                        continue
                    # Multiply by quantity for any per-unit
                    # price column (not totals/balances)
                    if is_per_unit and not is_already_total \
                            and qty_cols:
                        qty = safe_float(row.get(qty_cols[0])) \
                              or 1
                        v = v * qty
                    monthly[d.strftime('%Y-%m')] += v
                    print(f"TIME_SERIES | file={file_name} "
                          f"date_col={date_col} "
                          f"val_col={val_col} "
                          f"month={d.strftime('%Y-%m')} "
                          f"v={v} "
                          f"is_per_unit={is_per_unit} "
                          f"is_already_total={is_already_total}")

                if monthly:
                    sorted_months = sorted(monthly.items())
                    best = max(monthly, key=monthly.get)
                    worst = min(monthly, key=monthly.get)
                    time_series[val_col] = {
                        'by_month': {k: round(v, 2)
                                    for k, v in sorted_months},
                        'best_month': best,
                        'best_month_value': round(
                            monthly[best], 2),
                        'worst_month': worst,
                        'worst_month_value': round(
                            monthly[worst], 2),
                        'date_column_used': date_col,
                        'value_column_used': val_col,
                        'note': (
                            'multiply_by_qty_applied'
                            if (is_per_unit
                                and not is_already_total
                                and qty_cols)
                            else 'raw_column_values'
                        )
                    }
            if time_series:
                file_aggs['time_series'] = time_series

        # --- 5. Group performance: sum amounts by any group column ---
        if group_cols and (amount_cols or qty_cols):
            group_performance = {}
            for group_col in group_cols[:4]:
                for amount_col in amount_cols[:3]:
                    col_lower = str(amount_col).lower()
                    # Multiply by quantity for any per-unit
                    # price column (not totals/balances)
                    is_per_unit = any(h in col_lower
                        for h in ['unit_price', 'our_price',
                                  'list_price', 'sale_price',
                                  'unit_cost', 'cost_price'])
                    is_already_total = any(h in col_lower
                        for h in ['total', 'subtotal', 'balance',
                                  'amount_paid', 'balance_due'])
                    totals = defaultdict(float)
                    for row in rows:
                        grp = str(row.get(group_col, '') or '').strip()
                        if not grp:
                            continue
                        v = safe_float(row.get(amount_col))
                        if v is None:
                            continue
                        qty = 1
                        if is_per_unit and not is_already_total \
                                and qty_cols:
                            qty = safe_float(row.get(qty_cols[0])) \
                                  or 1
                            v = v * qty
                        totals[grp] += v
                        print(f"GROUP_PERF | file={file_name} "
                              f"group_col={group_col} "
                              f"amount_col={amount_col} "
                              f"grp={grp} v={v} "
                              f"is_per_unit={is_per_unit} "
                              f"qty_used={qty if qty_cols else 1}")
                    if totals:
                        ranked = sorted(totals.items(),
                                       key=lambda x: -x[1])
                        key = f'{amount_col}_by_{group_col}'
                        group_performance[key] = {
                            'ranked': {k: round(v, 2)
                                      for k, v in ranked[:20]},
                            'top': ranked[0][0],
                            'top_total': round(ranked[0][1], 2),
                            'bottom': ranked[-1][0],
                            'bottom_total': round(ranked[-1][1], 2)
                        }
            if group_performance:
                file_aggs['group_performance'] = group_performance

        # Build a query-ready summary of all
        # group_performance entries so the LLM
        # can match the right one to the question
        gp = file_aggs.get('group_performance', {})
        if gp:
            file_aggs['_rankings_index'] = {
                k: {
                    'grouped_by': k.split('_by_')[-1]
                        if '_by_' in k else k,
                    'measured_by': k.split('_by_')[0]
                        if '_by_' in k else k,
                    'top': v.get('top'),
                    'top_total': v.get('top_total'),
                    'bottom': v.get('bottom'),
                    'bottom_total': v.get('bottom_total'),
                    'all_ranked': v.get('ranked', {})
                }
                for k, v in gp.items()
            }

        # Build a query-ready summary of all
        # time_series entries so the LLM can
        # match the right one to the question
        ts = file_aggs.get('time_series', {})
        if ts:
            file_aggs['_timeseries_index'] = {
                k: {
                    'measures': k,
                    'date_column': v.get(
                        'date_column_used'),
                    'is_volume_adjusted': v.get(
                        'note') ==
                        'multiply_by_qty_applied',
                    'best_month': v.get('best_month'),
                    'best_month_value': v.get(
                        'best_month_value'),
                    'worst_month': v.get('worst_month'),
                    'worst_month_value': v.get(
                        'worst_month_value'),
                    'by_month': v.get('by_month', {})
                }
                for k, v in ts.items()
            }

        # Compute average rates per category
        # for columns that are percentages or rates
        # These should be averaged not summed
        rate_cols = [
            c for c in native_columns
            if any(h in str(c).lower()
                   for h in ['pct', 'rate', 'ratio',
                              'margin', 'percent',
                              'yield', 'efficiency'])
            and col_types.get(c) == 'amount'
        ]
        cat_cols_for_rates = [
            c for c in native_columns
            if col_types.get(c) in ('status', 'group')
        ]
        if rate_cols and cat_cols_for_rates:
            rate_averages = {}
            for rate_col in rate_cols[:3]:
                for cat_col in cat_cols_for_rates[:3]:
                    cat_avgs = defaultdict(list)
                    for row in rows:
                        cat = str(
                            row.get(cat_col, '') or ''
                        ).strip()
                        val = safe_float(
                            row.get(rate_col))
                        if cat and val is not None:
                            cat_avgs[cat].append(val)
                    if cat_avgs:
                        averaged = {
                            cat: round(
                                sum(vals)/len(vals), 2)
                            for cat, vals in
                            cat_avgs.items()
                        }
                        ranked = sorted(
                            averaged.items(),
                            key=lambda x: -x[1]
                        )
                        key = f'avg_{rate_col}_by_{cat_col}'
                        rate_averages[key] = {
                            'top': ranked[0][0],
                            'top_avg': ranked[0][1],
                            'all_averages': dict(ranked)
                        }
            if rate_averages:
                file_aggs['rate_averages'] = \
                    rate_averages

        aggregations[file_name] = file_aggs

    return aggregations


def _slim_aggregations(aggregations: dict) -> dict:
    """
    Reduce aggregations to only what the LLM needs
    to answer questions. Removes verbose by_month
    dicts and keeps only summary fields.
    This ensures all files fit within context.
    """
    slim = {}
    for file_name, data in aggregations.items():
        slim[file_name] = {}

        # Column types - keep as is, small
        if '_column_types' in data:
            slim[file_name]['column_types'] = \
                data['_column_types']

        # Numeric summaries - keep as is, small
        if 'numeric_summaries' in data:
            slim[file_name]['numeric_summaries'] = \
                data['numeric_summaries']

            # Add total revenue computed from
            # full group_performance not sample rows
            # This gives accurate totals regardless
            # of row sampling
            gp = data.get('group_performance', {})
            if gp:
                # Find the sale or revenue column
                # total by summing all group totals
                for gp_key, gp_val in gp.items():
                    measured = gp_key.split('_by_')[0] \
                        if '_by_' in gp_key else ''
                    if any(h in measured
                           for h in ['sale', 'revenue',
                                     'total_amount']):
                        all_vals = gp_val.get(
                            'ranked', {}).values()
                        if all_vals:
                            slim[file_name]\
                                ['total_revenue'] = {
                                'column': measured,
                                'total': round(
                                    sum(all_vals), 2),
                                'note': (
                                    'sum of all groups '
                                    'from full dataset'
                                )
                            }
                        break

        # Categorical breakdowns - keep top 5
        # groups only, drop value_counts
        if 'categorical_breakdowns' in data:
            slim_cat = {}
            for col, breakdown in \
                    data['categorical_breakdowns'].items():
                slim_cat[col] = {
                    'top_group': breakdown.get(
                        'top_group'),
                    'top_group_total': breakdown.get(
                        'top_group_total'),
                    'unique_count': breakdown.get(
                        'unique_count'),
                    'total_count': breakdown.get(
                        'total_count'),
                    'top_5_by_total': dict(
                        list(
                            breakdown.get(
                                'total_by_group',
                                {}).items()
                        )[:5]
                    )
                }
            slim[file_name]['categorical_breakdowns'] \
                = slim_cat

        # Conditional aggregations - keep as is,
        # already compact
        if 'conditional_aggregations' in data:
            slim[file_name]['conditional_aggregations']\
                = data['conditional_aggregations']

        # Rankings index - keep top and top_total
        # only, drop full ranked dict
        if '_rankings_index' in data:
            slim_rank = {}
            for k, v in data['_rankings_index'].items():
                slim_rank[k] = {
                    'grouped_by': v.get('grouped_by'),
                    'measured_by': v.get('measured_by'),
                    'top': v.get('top'),
                    'top_total': v.get('top_total'),
                    'top_5': dict(
                        list(
                            v.get('all_ranked',
                                  {}).items()
                        )[:5]
                    )
                }
            slim[file_name]['rankings'] = slim_rank

        # Time series index - drop by_month dict,
        # keep only best/worst summary
        if '_timeseries_index' in data:
            # Only include time series for files
            # that have at least one volume-adjusted
            # series. Files without quantity columns
            # (like invoices) only have raw billing
            # totals which should not be used for
            # sales period questions.
            has_volume_adjusted = any(
                v.get('is_volume_adjusted')
                for v in data['_timeseries_index']
                .values()
            )
            if has_volume_adjusted:
                slim_ts = {}
                for k, v in \
                        data['_timeseries_index']\
                        .items():
                    # Only include the volume
                    # adjusted series, skip raw ones
                    if not v.get('is_volume_adjusted'):
                        continue
                    slim_ts[k] = {
                        'measures': v.get('measures'),
                        'source_file': file_name,
                        'is_volume_adjusted': True,
                        'best_month': v.get(
                            'best_month'),
                        'best_month_value': v.get(
                            'best_month_value'),
                        'worst_month': v.get(
                            'worst_month'),
                        'worst_month_value': v.get(
                            'worst_month_value')
                    }
                if slim_ts:
                    slim[file_name]\
                        ['time_series_summary'] \
                        = slim_ts

        if 'rate_averages' in data:
            slim[file_name]['rate_averages'] = \
                data['rate_averages']

    return slim


class WorkspaceChatRequest(BaseModel):
    question: str


@router.post("/{workspace_id}/chat")
def workspace_chat(
    workspace_id: str,
    body: WorkspaceChatRequest
):
    try:
        from google import genai

        question = body.question

        # Load unified data
        conn = sqlite3.connect(_db_path())
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute("""
            SELECT file_name, unified_data
            FROM unified_datasets
            WHERE workspace_id = ?
        """, (workspace_id,))

        unified_files_full = []
        unified_files_slim = []
        for row in cursor.fetchall():
            data = json.loads(row['unified_data'])
            all_rows = data.get('rows', [])
            cols = data.get('columns', [])
            unified_files_full.append({
                'file': row['file_name'],
                'columns': cols,
                'rows': all_rows
            })
            unified_files_slim.append({
                'file': row['file_name'],
                'columns': cols,
                'rows': all_rows[:20]
            })

        cursor.execute("""
            SELECT file_1, column_1,
                   file_2, column_2
            FROM cross_file_relationships
            WHERE workspace_id = ?
            AND confirmed = 1
        """, (workspace_id,))

        relationships = []
        for r in cursor.fetchall():
            relationships.append(
                f"{r['file_1']}.{r['column_1']} "
                f"= {r['file_2']}.{r['column_2']}"
            )

        conn.close()

        if not unified_files_full:
            raise HTTPException(
                status_code=400,
                detail="No unified data found. Run unification first."
            )

        # Build evidence
        aggregations = compute_aggregations(unified_files_full)
        aggregations = _slim_aggregations(aggregations)

        # Build file metadata with key column values
        # for orphan detection across files.
        # Only include ID and name columns not all
        # rows - keeps evidence compact while
        # enabling cross-file key comparison.
        known_prefixes = tuple(
            f['file'].replace('.csv', '')
            .replace('.xlsx', '') + '_'
            for f in unified_files_slim
        )

        def get_key_values(rows, columns):
            id_cols = [
                c for c in columns
                if any(h in str(c).lower()
                       for h in ['_id', 'id_',
                                 'code', 'number'])
                and not str(c).startswith(
                    tuple(known_prefixes)
                )
            ]
            name_cols = [
                c for c in columns
                if any(h in str(c).lower()
                       for h in ['name', 'company',
                                 'customer'])
                and not str(c).startswith(
                    tuple(known_prefixes)
                )
            ]
            key_cols = (id_cols + name_cols)[:3]
            return {
                col: list(dict.fromkeys(
                    str(r.get(col, '')).strip()
                    for r in rows
                    if r.get(col)
                ))
                for col in key_cols
            }

        file_metadata = []
        for i, f in enumerate(unified_files_slim):
            full_rows = unified_files_full[i]['rows']
            cols = f['columns']
            native_cols = [
                c for c in cols
                if not str(c).startswith(
                    tuple(known_prefixes)
                )
            ]
            file_metadata.append({
                'file': f['file'],
                'columns': native_cols,
                'total_rows': len(full_rows),
                'key_values': get_key_values(
                    full_rows, native_cols
                )
            })

        # Pre-compute cross-file orphan analysis
        # so LLM does not need to compare lists
        # This finds IDs in one file missing from
        # another using confirmed relationships
        def is_master_file(filename):
            f = filename.lower()
            return any(h in f for h in [
                'customer', 'product', 'employee',
                'vendor', 'supplier', 'item',
                'catalog', 'master', 'reference',
                'contact', 'account'
            ])

        def is_transaction_file(filename):
            f = filename.lower()
            return any(h in f for h in [
                'order', 'invoice', 'sale',
                'transaction', 'purchase',
                'shipment', 'payment', 'ledger'
            ])

        orphan_analysis = []
        for rel_str in relationships:
            # rel_str format: file1.col1 = file2.col2
            try:
                left, right = rel_str.split(' = ')
                file1, col1 = left.rsplit('.', 1)
                file2, col2 = right.rsplit('.', 1)

                # Only check orphans between a
                # transaction file and a master file
                # Transaction vs transaction orphans
                # are expected and not meaningful
                # e.g. orders without invoices is
                # normal business - not every order
                # has been invoiced yet
                both_transaction = (
                    is_transaction_file(file1) and
                    is_transaction_file(file2)
                )
                if both_transaction:
                    continue

                # Skip name column comparisons
                # Names are unreliable join keys
                # because same entity can have
                # different formats across files
                name_hints = ['name', 'company',
                              'description', 'title']
                col1_is_name = any(
                    h in col1.lower()
                    for h in name_hints
                )
                col2_is_name = any(
                    h in col2.lower()
                    for h in name_hints
                )
                if col1_is_name and col2_is_name:
                    continue

                # Only compare ID-type columns
                id_hints = ['_id', 'id_', 'order_id',
                            'invoice_id', 'product_id',
                            'customer_id', 'code',
                            'number', 'ref', 'key']
                col1_is_id = any(
                    h in col1.lower()
                    for h in id_hints
                )
                col2_is_id = any(
                    h in col2.lower()
                    for h in id_hints
                )
                if not col1_is_id or not col2_is_id:
                    continue

                # Get values from full data (normalised)
                vals1 = set()
                vals2 = set()
                for f in unified_files_full:
                    if f['file'] == file1:
                        vals1 = set(
                            str(r.get(col1, '')).strip().lower()
                            for r in f['rows']
                            if r.get(col1)
                        )
                    if f['file'] == file2:
                        vals2 = set(
                            str(r.get(col2, '')).strip().lower()
                            for r in f['rows']
                            if r.get(col2)
                        )

                if vals1 and vals2:
                    in_1_not_2 = sorted(vals1 - vals2)
                    in_2_not_1 = sorted(vals2 - vals1)
                    orphan_analysis.append({
                        'relationship': rel_str,
                        'file1': file1,
                        'col1': col1,
                        'file2': file2,
                        'col2': col2,
                        f'in_{file1}_not_in_{file2}':
                            in_1_not_2,
                        f'in_{file2}_not_in_{file1}':
                            in_2_not_1,
                        'orphans_found':
                            len(in_1_not_2) > 0 or
                            len(in_2_not_1) > 0
                    })
            except Exception:
                pass

        # Also detect name mismatches for same ID
        # across files - same ID different name
        # is a data quality issue not a true orphan
        name_mismatches = []
        for f1 in unified_files_full:
            for f2 in unified_files_full:
                if f1['file'] >= f2['file']:
                    continue
                # Find shared ID columns
                cols1 = set(f1['columns'])
                cols2 = set(f2['columns'])
                id_cols = [
                    c for c in cols1 & cols2
                    if any(h in str(c).lower()
                           for h in ['customer_id',
                                     'product_id',
                                     'order_id'])
                ]
                for id_col in id_cols[:2]:
                    # Build id->name maps
                    name_col_hints = [
                        'name', 'company', 'customer'
                    ]
                    name_col1 = next((
                        c for c in f1['columns']
                        if any(h in str(c).lower()
                               for h in name_col_hints)
                        and c != id_col
                    ), None)
                    name_col2 = next((
                        c for c in f2['columns']
                        if any(h in str(c).lower()
                               for h in name_col_hints)
                        and c != id_col
                    ), None)
                    if not name_col1 or not name_col2:
                        continue
                    map1 = {
                        str(r.get(id_col, '')).strip():
                        str(r.get(name_col1, '')).strip()
                        for r in f1['rows']
                        if r.get(id_col)
                    }
                    map2 = {
                        str(r.get(id_col, '')).strip():
                        str(r.get(name_col2, '')).strip()
                        for r in f2['rows']
                        if r.get(id_col)
                    }
                    for id_val in set(
                        map1.keys()) & set(map2.keys()
                    ):
                        n1 = map1[id_val]
                        n2 = map2[id_val]
                        id_looks_like_code = bool(
                            re.match(
                                r'^[a-zA-Z]\d+$',
                                str(id_val or '').strip()
                            )
                        )
                        if n1.lower() != n2.lower() \
                                and id_looks_like_code:
                            name_mismatches.append({
                                'id': id_val,
                                'id_column': id_col,
                                f'name_in_{f1["file"]}':
                                    n1,
                                f'name_in_{f2["file"]}':
                                    n2
                            })

        # Add to evidence
        evidence = {
            "pre_computed_aggregations": aggregations,
            "confirmed_relationships": relationships,
            "orphan_analysis": orphan_analysis,
            "name_mismatches": name_mismatches,
            "file_metadata": file_metadata
        }

        system_prompt = """You are a senior data analyst.
You answer questions using pre_computed_aggregations.
You do not have access to raw data rows.
All answers must come from the aggregations.

AGGREGATION STRUCTURE:
Each file in pre_computed_aggregations contains:

rankings
  Every group_by combination pre-computed.
  Each entry has: grouped_by, measured_by,
  top, top_total, top_5 dict.
  For performance questions find the entry
  where grouped_by matches what the question
  asks about (sales_rep, category, region etc)
  and measured_by represents the right metric
  (prefer sale_price over unit_price over cost).

time_series_summary
  Monthly totals for each numeric column.
  Each entry has: measures, is_volume_adjusted,
  best_month, best_month_value, worst_month,
  worst_month_value.
  For sales questions use the entry where
  is_volume_adjusted=true from the transactions
  file. Never use invoice date-based series
  for sales period questions.

categorical_breakdowns
  Counts and totals by category columns.
  Each entry has: top_group, top_group_total,
  top_5_by_total, unique_count, total_count.
  Use for breakdowns by type or segment.

conditional_aggregations
  Totals split by status values.
  Format: amount_by_status.
  Each status has sum and count.
  Use for outstanding or filtered amounts.

numeric_summaries
  Overall sum, min, max, avg, count per column.
  Use for total figures across all records.

file_metadata
  Shows file names, columns, total row counts, and
  key_values (unique ID and name values per column).
  Use key_values for orphan detection across files.
  Use total_rows for counts. Do not use for financial
  figures.

confirmed_relationships
  How files relate to each other via join keys.
  Use to trace cross-file questions.

STRICT RULES:
1. Every number in your answer must come from
   aggregations. Never calculate yourself.
   Never estimate. Never use file_metadata
   row counts for financial figures.

2. For rep/agent/person performance always use
   rankings where grouped_by contains the person
   column. Pick measured_by that represents
   revenue (sale, revenue, total over cost/unit).
   Read top and top_total directly.

3. For time/period and monthly sales questions
     use time_series_summary. Find the entry where
     is_volume_adjusted=true. This is always the
     correct source for sales questions because it
     means quantity times price was applied making
     it true transaction revenue. If multiple files
     have is_volume_adjusted=true entries, use the
     one whose source_file is not an invoice or
     billing file. Read best_month and
     best_month_value directly from that entry.

4. For margin or profitability by category use
     rate_averages from the products file. Find
     the entry whose key starts with avg_ and
     contains margin. Read top and top_avg to
     get the highest margin category and its
     average percentage. Report it as a percentage.

5. For overdue or outstanding balance questions
     use conditional_aggregations. Find keys that
     start with outstanding_ - these contain
     pre-filtered balances for unpaid records only.
     Read by_entity for the full customer list and
     total_outstanding for the overall total.
     Never use total_amount or subtotal for this
     question - only balance_due or amount_due
     columns represent what is actually owed.

6. For questions about whether records in one
     file exist in another file use orphan_analysis.
     Each entry shows a confirmed relationship and
     the IDs missing from each side.
     
     IMPORTANT: orphan_analysis only compares
     transaction files against master/reference
     files. It does not compare transaction files
     against each other because that is expected
     behaviour (not every order has an invoice yet).
     
     For the specific question "do customers in
     orders exist in customer master" look for the
     orphan_analysis entry where one file is the
     orders file and the other is the customers
     file. If in_orders_not_in_customers is empty
     then all customers in orders exist in the
     master. Also check name_mismatches for cases
     where the same customer_id has a different
     name spelling across files and report those
     as data quality issues.

7. For executive summary questions build the
     answer using only these sources:
     - Total revenue: use total_revenue.total from
       the transactions file if available. This is
       computed from the full dataset not a sample.
       If not available use numeric_summaries sum
       for the sale or revenue column.
     - Top rep: rankings sale_price_by_sales_rep
       top and top_total
     - Top month: time_series_summary best_month
       and best_month_value
     - Top category by margin: rate_averages top
     - Outstanding balance: conditional_aggregations
       total_outstanding from outstanding_ key
     - Customer count: file_metadata total_rows
       for the customers file
     - Top customer by lifetime value: rankings
       total_lifetime_value_by_company_name top
     Never mix sources for the same metric.
     Never report the same number for different
     metrics such as using rep total for region
     total.

8. Never show raw IDs. Always use name fields.
9. Never say based on the data.
10. Cite source files at end only.

RESPONSE FORMAT:
- Direct answer on first line
- Bullets for 3 or more items
- Under 150 words unless full list needed
- Source at end only

Be direct. Be accurate. Be brief."""

        user_message = f"""Unified dataset:

{json.dumps(evidence, indent=2)}

Question: {question}"""

        # Call Gemini
        api_key = os.getenv('GEMINI_API_KEY')
        if not api_key:
            raise HTTPException(
                status_code=500,
                detail="GEMINI_API_KEY not set"
            )

        client = genai.Client(api_key=api_key)

        print("=== AGGREGATION KEYS SENT TO GEMINI ===")
        for fname, agg in aggregations.items():
            print(f"  {fname}: {list(agg.keys())}")
        print("=== EVIDENCE SENT TO GEMINI (first 6000 chars) ===")
        import json as _json
        print(_json.dumps(evidence, indent=2)[:6000])
        print("=== END EVIDENCE ===")

        print("=== SLIM AGG KEYS PER FILE ===")
        for fname, fdata in aggregations.items():
            print(f"\n{fname}:")
            for k, v in fdata.items():
                if k == 'rankings':
                    for rk, rv in v.items():
                        print(f"  rankings.{rk}.top = {rv.get('top')} / {rv.get('top_total')}")
                elif k == 'time_series_summary':
                    for tk, tv in v.items():
                        print(f"  timeseries.{tk}.best = {tv.get('best_month')} / {tv.get('best_month_value')} vol_adj={tv.get('is_volume_adjusted')}")
                else:
                    print(f"  {k}: present")
        print("=== END SLIM AGG KEYS ===")

        response = client.models.generate_content(
            model='gemini-2.5-flash-lite',
            config={'system_instruction': system_prompt},
            contents=user_message
        )

        return {
            "answer": response.text,
            "files_used": [f['file'] for f in unified_files_slim],
            "workspace_id": workspace_id
        }

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Chat error: {str(e)}"
        )
