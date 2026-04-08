import sqlite3
import json
import re
from datetime import datetime


def run_unification(workspace_id, db_path):
    """
    Main function. Runs after human confirms
    relationships in the ER diagram.

    Steps:
    1. Load all files and confirmed relationships
    2. Detect and remove duplicates
    3. Standardize formats
    4. Enrich records using confirmed connections
    5. Store unified data in SQLite
    6. Return unification report
    """

    report = {
        'workspace_id': workspace_id,
        'files_processed': 0,
        'duplicates_removed': 0,
        'columns_enriched': 0,
        'format_fixes': 0,
        'unified_tables': [],
        'errors': []
    }

    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # Load all files in workspace
        cursor.execute("""
            SELECT wf.workbook_id, wf.file_name,
                   w.parsed_json
            FROM workspace_files wf
            JOIN workbooks w
              ON wf.workbook_id = w.workbook_id
            WHERE wf.workspace_id = ?
        """, (workspace_id,))

        files = {}
        for row in cursor.fetchall():
            parsed = json.loads(
                row['parsed_json'] or '{}'
            )
            sheets = parsed.get('sheets', [])
            main_sheet = None
            for sheet in sheets:
                if sheet.get('row_count', 0) > 0:
                    main_sheet = sheet
                    break

            if main_sheet:
                files[row['file_name']] = {
                    'workbook_id': row['workbook_id'],
                    'file_name': row['file_name'],
                    'columns': main_sheet.get('columns', []),
                    'rows': main_sheet.get('sample_rows', [])
                }

        # Load confirmed relationships
        cursor.execute("""
            SELECT file_1, column_1,
                   file_2, column_2,
                   relationship_type
            FROM cross_file_relationships
            WHERE workspace_id = ?
            AND confirmed = 1
            AND rejected = 0
        """, (workspace_id,))

        relationships = [
            dict(r) for r in cursor.fetchall()
        ]

        conn.close()

    except Exception as e:
        report['errors'].append(str(e))
        return report

    # Process each file
    unified_data = {}

    for file_name, file_data in files.items():
        rows = file_data['rows']
        columns = file_data['columns']

        if not rows:
            continue

        # Step 1: Remove duplicates
        original_count = len(rows)
        rows = deduplicate_rows(rows)
        dupes_removed = original_count - len(rows)
        report['duplicates_removed'] += dupes_removed

        # Step 2: Standardize formats
        rows, fixes = standardize_formats(rows, columns)
        report['format_fixes'] += fixes

        unified_data[file_name] = {
            'columns': columns,
            'rows': rows,
            'original_count': original_count,
            'clean_count': len(rows),
            'dupes_removed': dupes_removed
        }

        report['files_processed'] += 1

    # Step 3: Enrich records using relationships
    original_column_counts = {
        fn: len(d['columns']) for fn, d in unified_data.items()
    }
    for rel in relationships:
        file1 = rel['file_1']
        file2 = rel['file_2']
        col1 = rel['column_1']
        col2 = rel['column_2']

        if file1 not in unified_data:
            continue
        if file2 not in unified_data:
            continue

        # Build lookup from file1
        lookup = {}
        for row in unified_data[file1]['rows']:
            key = str(row.get(col1, '')).strip().lower()
            if key:
                lookup[key] = row

        # Enrich file2 rows with file1 data
        for row in unified_data[file2]['rows']:
            val = str(row.get(col2, '')).strip().lower()
            if val and val in lookup:
                source_row = lookup[val]
                # Add columns from file1 that
                # don't already exist in file2
                for src_col, src_val in source_row.items():
                    enriched_key = f"{file1.replace('.csv','').replace('.xlsx','')}_{src_col}"
                    if enriched_key not in row:
                        row[enriched_key] = src_val

        # Update column list for file2
        if unified_data[file2]['rows']:
            unified_data[file2]['columns'] = list(
                unified_data[file2]['rows'][0].keys()
            )

    # Count new cross-file columns added per file
    for fn, d in unified_data.items():
        orig = original_column_counts.get(fn, 0)
        curr = len(d['columns'])
        report['columns_enriched'] += max(0, curr - orig)

    # Step 4: Detect orphan records
    orphans = detect_orphans(
        unified_data, relationships
    )

    # Step 5: Save to SQLite
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # Create unified_datasets table if not exists
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS
            unified_datasets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                workspace_id TEXT NOT NULL,
                file_name TEXT NOT NULL,
                unified_data TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)

        # Delete existing unified data
        cursor.execute("""
            DELETE FROM unified_datasets
            WHERE workspace_id = ?
        """, (workspace_id,))

        # Insert unified data per file
        for file_name, data in unified_data.items():
            cursor.execute("""
                INSERT INTO unified_datasets
                (workspace_id, file_name,
                 unified_data, created_at)
                VALUES (?, ?, ?, ?)
            """, (
                workspace_id,
                file_name,
                json.dumps(data),
                datetime.now().isoformat()
            ))

            report['unified_tables'].append({
                'file': file_name,
                'clean_rows': data['clean_count'],
                'columns': len(data['columns']),
                'dupes_removed': data['dupes_removed']
            })

        conn.commit()
        conn.close()

    except Exception as e:
        report['errors'].append(
            f"Save error: {str(e)}"
        )

    report['orphan_records'] = orphans
    report['status'] = 'complete'

    return report


def deduplicate_rows(rows):
    """
    Removes completely identical rows.
    Keeps first occurrence.
    """
    seen = set()
    unique_rows = []

    for row in rows:
        # Create a hashable key from row values
        row_key = tuple(
            sorted(
                (k, str(v))
                for k, v in row.items()
                if v is not None and str(v).strip()
            )
        )

        if row_key not in seen:
            seen.add(row_key)
            unique_rows.append(row)

    return unique_rows


def standardize_formats(rows, columns):
    """
    Standardizes common format issues:
    - Date formats → YYYY-MM-DD
    - Phone numbers → consistent format
    - Currency strings → numbers
    - Whitespace cleanup
    """
    fixes = 0

    # Date patterns to detect and fix
    date_patterns = [
        # MM/DD/YYYY → YYYY-MM-DD
        (re.compile(r'^(\d{1,2})/(\d{1,2})/(\d{4})$'),
         lambda m: f"{m.group(3)}-{m.group(1).zfill(2)}-{m.group(2).zfill(2)}"),
        # DD-MM-YYYY → YYYY-MM-DD
        (re.compile(r'^(\d{1,2})-(\d{1,2})-(\d{4})$'),
         lambda m: f"{m.group(3)}-{m.group(2).zfill(2)}-{m.group(1).zfill(2)}"),
        # MM-DD-YYYY → YYYY-MM-DD
        (re.compile(r'^(\d{1,2})-(\d{1,2})-(\d{4})$'),
         lambda m: f"{m.group(3)}-{m.group(1).zfill(2)}-{m.group(2).zfill(2)}"),
    ]

    for row in rows:
        for col, val in list(row.items()):
            if val is None or val == '':
                continue

            str_val = str(val).strip()

            # Fix dates
            for pattern, formatter in date_patterns:
                match = pattern.match(str_val)
                if match:
                    new_val = formatter(match)
                    if new_val != str_val:
                        row[col] = new_val
                        fixes += 1
                    break

            # Fix currency strings
            # "$50,000" → 50000
            if isinstance(val, str):
                currency_match = re.match(
                    r'^\$?([\d,]+\.?\d*)$',
                    str_val.replace(',', '')
                )
                if currency_match:
                    try:
                        row[col] = float(
                            currency_match.group(1)
                            .replace(',', '')
                        )
                        if row[col] != val:
                            fixes += 1
                    except ValueError:
                        pass

            # Trim whitespace
            if isinstance(val, str) and \
               val != val.strip():
                row[col] = val.strip()
                fixes += 1

    return rows, fixes


def detect_orphans(unified_data, relationships):
    """
    Finds records in child tables that have
    no matching record in the parent table.
    These are data integrity issues.
    """
    from collections import defaultdict

    def normalize(val):
        return str(val).strip().lower()

    def is_name_col(col):
        col_lower = str(col).lower()
        return any(h in col_lower for h in [
            'name', 'company', 'description',
            'title', 'label'
        ])

    orphans = []

    for rel in relationships:
        file1 = rel['file_1']
        file2 = rel['file_2']
        col1 = rel['column_1']
        col2 = rel['column_2']

        if file1 not in unified_data:
            continue
        if file2 not in unified_data:
            continue

        # Skip if both columns are name-type columns
        # Name columns are unreliable for orphan detection
        # because the same entity can have different formats
        if is_name_col(col1) and is_name_col(col2):
            continue

        # Build set of parent values (normalised)
        other_values = {
            normalize(row.get(col1, ''))
            for row in unified_data[file1]['rows']
            if normalize(row.get(col1, ''))
        }

        # Find child records with no parent
        for row in unified_data[file2]['rows']:
            value = row.get(col2)
            if value is None or not normalize(value):
                continue
            if normalize(value) not in other_values:
                orphans.append({
                    'file': file2,
                    'column': col2,
                    'value': str(value),
                    'issue': f"'{value}' in {file2}.{col2} has no match in {file1}.{col1}"
                })

    # Deduplicate by unique value per file/column
    seen = set()
    unique_orphans = []
    for o in orphans:
        key = o.get('issue', str(o))
        if key not in seen:
            seen.add(key)
            unique_orphans.append(o)
    orphans = unique_orphans

    # Group orphans by pattern and summarize
    grouped = defaultdict(list)
    for orphan in orphans:
        issue = orphan.get('issue', '')
        parts = issue.split(' in ', 1)
        if len(parts) > 1:
            key = 'in ' + parts[1]
        else:
            key = issue
        grouped[key].append(orphan)

    summarized = []
    for pattern, items in grouped.items():
        if len(items) == 1:
            summarized.append(items[0])
        else:
            first_issue = items[0].get('issue', '')
            val_part = first_issue.split(' in ')[0]
            summarized.append({
                'issue': (
                    f"{len(items)} values in "
                    f"{pattern.replace('in ', '', 1)}"
                    f" (e.g. {val_part})"
                )
            })

    orphans = summarized

    return orphans


def get_unified_data(workspace_id, file_name, db_path):
    """
    Retrieves unified data for a specific file.
    Used by chat to query clean data.
    """
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute("""
            SELECT unified_data
            FROM unified_datasets
            WHERE workspace_id = ?
            AND file_name = ?
        """, (workspace_id, file_name))

        row = cursor.fetchone()
        conn.close()

        if row:
            return json.loads(row['unified_data'])
        return None

    except Exception:
        return None
