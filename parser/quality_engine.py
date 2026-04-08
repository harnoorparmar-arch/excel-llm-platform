"""Generic quality checks for any tabular data — financial, sales, HR, inventory, medical, etc."""
from collections import Counter
from typing import Any

ERROR_VALUES = {
    '#DIV/0!', '#N/A', '#NAME?',
    '#NULL!', '#NUM!', '#REF!',
    '#VALUE!', '#ERROR!',
    'DIV/0!', 'N/A', 'NAME?',
    'NULL!', 'NUM!', 'REF!', 'VALUE!'
}

SUBTOTAL_KEYWORDS = {
    'total', 'subtotal', 'sub-total',
    'sub total', 'grand total',
    'sum', 'aggregate', 'net total',
    'overall', 'combined', 'totals'
}


def _is_single_letter_column(col_name: str) -> bool:
    """Return True if column name is a single letter (A, B, C, etc)."""
    s = str(col_name).strip()
    return len(s) == 1 and s.isalpha()


def _is_numeric(val: Any) -> bool:
    """Return True if value is numeric (int, float, or numeric string)."""
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return True
    if val is None or val == "":
        return False
    s = str(val).strip()
    if not s:
        return False
    try:
        float(s.replace(",", "").replace("%", ""))
        return True
    except ValueError:
        return False


def _non_empty_count(row: dict[str, Any], columns: list[str]) -> int:
    """Count non-empty cells in a row."""
    return sum(1 for c in columns if row.get(c) is not None and row.get(c) != "")


def run_quality_checks(
    workbook_id: str,
    sheets_data: list[dict[str, Any]],
    schema_mappings: list[dict[str, Any]],
) -> dict[str, Any]:
    """
    Run generic structural quality checks across all sheets.
    Works on any tabular data — no domain assumptions.
    schema_mappings is retained for API compatibility but not used.
    """
    issues: list[dict[str, Any]] = []

    for sheet in sheets_data:
        sheet_name = sheet.get("sheet_name", "")
        columns = sheet.get("columns") or []
        sample_rows = sheet.get("sample_rows") or []
        row_count = len(sample_rows)

        # Rule 5 - empty_sheet
        if row_count == 0:
            issues.append({
                "severity": "low",
                "rule": "empty_sheet",
                "sheet": sheet_name,
                "column": None,
                "message": "Sheet has 0 rows",
            })
            continue

        if not columns:
            continue

        # Rule 1 - null_spike (skip if < 5 rows)
        if row_count >= 5:
            for col in columns:
                if _is_single_letter_column(col):
                    continue
                null_count = sum(
                    1 for row in sample_rows
                    if row.get(col) is None or row.get(col) == ""
                )
                pct = null_count / row_count
                if pct > 0.85:
                    issues.append({
                        "severity": "high",
                        "rule": "null_spike",
                        "sheet": sheet_name,
                        "column": col,
                        "message": f"{int(pct * 100)}% of values are empty in {col} column",
                    })
                elif pct > 0.60:
                    issues.append({
                        "severity": "medium",
                        "rule": "null_spike",
                        "sheet": sheet_name,
                        "column": col,
                        "message": f"{int(pct * 100)}% of values are empty in {col} column",
                    })

        # Rule 2 - duplicate_rows (only count rows with >= 3 non-empty cells)
        row_tuples = [tuple(row.get(c) for c in columns) for row in sample_rows]
        seen: dict[tuple, int] = {}
        for i, rt in enumerate(row_tuples):
            if _non_empty_count(sample_rows[i], columns) >= 3:
                seen[rt] = seen.get(rt, 0) + 1
        dup_sets = sum(1 for c in seen.values() if c > 1)
        if dup_sets > 0:
            issues.append({
                "severity": "medium",
                "rule": "duplicate_rows",
                "sheet": sheet_name,
                "column": None,
                "message": f"Sheet has {dup_sets} sets of identical duplicate rows",
            })

        # Rule 3 - type_inconsistency
        for col in columns:
            non_empty = [row.get(col) for row in sample_rows if row.get(col) is not None and row.get(col) != ""]
            total = len(non_empty)
            if total < 3:
                continue
            numeric_count = sum(1 for v in non_empty if _is_numeric(v))
            non_numeric_count = total - numeric_count
            pct_numeric = numeric_count / total
            pct_non_numeric = non_numeric_count / total
            if pct_numeric > 0.70 and pct_non_numeric > 0.20:
                issues.append({
                    "severity": "medium",
                    "rule": "type_inconsistency",
                    "sheet": sheet_name,
                    "column": col,
                    "message": f"Column {col} has mixed types: {numeric_count}/{total} numeric, {non_numeric_count} non-numeric",
                })

        # Rule 4 - single_value_column
        if row_count > 5:
            for col in columns:
                unique_vals = set()
                for row in sample_rows:
                    v = row.get(col)
                    if v is not None and v != "":
                        unique_vals.add(str(v).strip())
                if len(unique_vals) == 1 and len(unique_vals) > 0:
                    issues.append({
                        "severity": "low",
                        "rule": "single_value_column",
                        "sheet": sheet_name,
                        "column": col,
                        "message": f"Column {col} has only 1 unique value across {row_count} rows",
                    })

        # Rule 6 - header_in_data (one issue per sheet)
        header_set = {str(c).strip() for c in columns}
        for row in sample_rows:
            for col in columns:
                val = row.get(col)
                if val is not None and val != "" and str(val).strip() in header_set:
                    issues.append({
                        "severity": "low",
                        "rule": "header_in_data",
                        "sheet": sheet_name,
                        "column": col,
                        "message": f"Data row contains header-like value '{val}' in column {col}",
                    })
                    break
            else:
                continue
            break

        # Rule 7 - formula_errors
        # Use pre-computed formula_error_cells from sheet metadata (e.g. CSV with na_values),
        # otherwise scan sample_rows for error values
        error_cells = sheet.get("formula_error_cells") or []
        if not error_cells:
            for row_idx, row in enumerate(sample_rows):
                for col in columns:
                    val = row.get(col)
                    if val is not None and str(val).strip() in ERROR_VALUES:
                        error_cells.append({
                            "column": col,
                            "row": row_idx,
                            "error": str(val).strip()
                        })

        if len(error_cells) > 0:
            error_types = Counter(e['error'] for e in error_cells)
            severity = 'high' if len(error_cells) > 3 else 'medium'
            issues.append({
                'severity': severity,
                'rule': 'formula_errors',
                'sheet': sheet_name,
                'column': 'multiple',
                'message': (
                    f"{len(error_cells)} formula "
                    f"error(s) found: "
                    f"{dict(error_types)}. "
                    f"Check for broken formulas "
                    f"or missing references."
                )
            })

        # Rule 8 - subtotal_rows
        subtotal_rows = []
        label_col = None
        for col in columns:
            non_empty = [row.get(col) for row in sample_rows if row.get(col) is not None and row.get(col) != ""]
            if not non_empty:
                continue
            string_vals = [v for v in non_empty if isinstance(v, str)]
            if len(string_vals) > len(non_empty) * 0.5:
                label_col = col
                break

        if label_col:
            for row_idx, row in enumerate(sample_rows):
                val = row.get(label_col)
                if val is None:
                    continue
                val_str = str(val).strip().lower()
                is_subtotal = any(kw in val_str for kw in SUBTOTAL_KEYWORDS)
                if is_subtotal:
                    subtotal_rows.append({
                        'row': row_idx,
                        'label': str(val).strip()
                    })

        if len(subtotal_rows) > 0:
            issues.append({
                'severity': 'low',
                'rule': 'subtotal_rows',
                'sheet': sheet_name,
                'column': label_col or 'unknown',
                'message': (
                    f"{len(subtotal_rows)} subtotal/"
                    f"total row(s) detected: "
                    f"{[r['label'] for r in subtotal_rows[:3]]}. "
                    f"These rows may cause double "
                    f"counting in aggregations."
                ),
                'subtotal_row_indices': [r['row'] for r in subtotal_rows]
            })

        # Rule 9 - multiple_tables_on_sheet
        if sheet.get('has_multiple_tables'):
            table_count = sheet.get('table_count', 2)
            issues.append({
                'severity': 'medium',
                'rule': 'multiple_tables_on_sheet',
                'sheet': sheet_name,
                'column': 'N/A',
                'message': (
                    f"Sheet contains {table_count} separate "
                    f"tables. Only the largest is being "
                    f"analyzed. Other tables may contain "
                    f"additional data."
                )
            })

    # Scoring
    high_count = sum(1 for i in issues if i["severity"] == "high")
    medium_count = sum(1 for i in issues if i["severity"] == "medium")
    low_count = sum(1 for i in issues if i["severity"] == "low")
    high_count = int(high_count)
    medium_count = int(medium_count)
    low_count = int(low_count)
    print(f"[quality_engine] high_count={high_count} (type={type(high_count).__name__}), medium_count={medium_count}, low_count={low_count}")
    score = 1.0 - (high_count * 0.05) - (medium_count * 0.02) - (low_count * 0.005)
    overall_score = round(max(0.0, min(1.0, score)), 2)
    print(f"[quality_engine] Score calculated: {score}, overall_score={overall_score} (type={type(overall_score).__name__})")

    return {
        "workbook_id": workbook_id,
        "overall_score": overall_score,
        "total_issues": len(issues),
        "issues": issues,
    }
