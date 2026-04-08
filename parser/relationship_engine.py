import json
import os
import sqlite3
from difflib import SequenceMatcher
from uuid import uuid4


def normalize_column_names_with_llm(files_data):
    """
    Disabled - returns input unchanged.
    Value overlap is sufficient for matching.
    LLM normalization caused messy column names.
    """
    for file_data in files_data:
        # Just use original column names as-is
        file_data['normalized_columns'] = [
            str(col).lower().strip().replace(' ', '_')
            for col in file_data.get('columns', [])
        ]
    return files_data


def detect_cross_file_relationships(workspace_id, db_path):
    """
    Detects potential relationships between
    columns across all files in a workspace.
    Returns list of potential relationships
    and primary key candidates.
    """

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

        files = []
        for row in cursor.fetchall():
            parsed = json.loads(row['parsed_json'] or '{}')
            sheets = parsed.get('sheets', [])

            # For CSV files there is usually
            # one sheet. Get first sheet with data.
            main_sheet = None
            for sheet in sheets:
                if sheet.get('row_count', 0) > 0:
                    main_sheet = sheet
                    break

            if main_sheet:
                files.append({
                    'workbook_id': row['workbook_id'],
                    'file_name': row['file_name'],
                    'columns': main_sheet.get('columns', []),
                    'sample_rows': main_sheet.get('sample_rows', [])
                })

        conn.close()

    except Exception as e:
        return {
            'relationships': [],
            'primary_keys': [],
            'error': str(e)
        }

    if len(files) < 2:
        return {
            'relationships': [],
            'primary_keys': detect_primary_keys(files),
            'message': 'Need at least 2 files to detect relationships'
        }

    # Normalize column names with LLM for better matching
    print("Normalizing column names with LLM...")
    files = normalize_column_names_with_llm(files)

    relationships = []

    # Compare every pair of files
    for i in range(len(files)):
        for j in range(i + 1, len(files)):
            file1 = files[i]
            file2 = files[j]

            cols1 = file1['columns']
            cols2 = file2['columns']
            norm_cols1 = file1.get('normalized_columns', cols1)
            norm_cols2 = file2.get('normalized_columns', cols2)

            for idx1, col1 in enumerate(cols1):
                if len(str(col1).strip()) <= 1:
                    continue

                norm1 = (
                    norm_cols1[idx1]
                    if idx1 < len(norm_cols1)
                    else str(col1).lower()
                )

                for idx2, col2 in enumerate(cols2):
                    if len(str(col2).strip()) <= 1:
                        continue

                    norm2 = (
                        norm_cols2[idx2]
                        if idx2 < len(norm_cols2)
                        else str(col2).lower()
                    )

                    # Use normalized names for name similarity
                    name_sim = SequenceMatcher(
                        None,
                        str(norm1).lower().strip(),
                        str(norm2).lower().strip()
                    ).ratio()

                    # Use original column names for value lookup
                    vals1 = set()
                    for row in file1['sample_rows']:
                        v = row.get(col1, '')
                        if v and str(v).strip():
                            vals1.add(str(v).strip().lower())

                    vals2 = set()
                    for row in file2['sample_rows']:
                        v = row.get(col2, '')
                        if v and str(v).strip():
                            vals2.add(str(v).strip().lower())

                    if not vals1 or not vals2:
                        continue

                    overlap = len(vals1 & vals2)
                    overlap_ratio = overlap / min(len(vals1), len(vals2))

                    confidence = name_sim * 0.3 + overlap_ratio * 0.7

                    if confidence < 0.25:
                        continue

                    unique1 = len(vals1)
                    unique2 = len(vals2)

                    if unique1 <= unique2:
                        rel_type = 'one_to_many'
                        pk_side = file1['file_name']
                        pk_col = col1
                        fk_side = file2['file_name']
                        fk_col = col2
                    else:
                        rel_type = 'one_to_many'
                        pk_side = file2['file_name']
                        pk_col = col2
                        fk_side = file1['file_name']
                        fk_col = col1

                    # Use ORIGINAL column names (from file columns) for DB storage
                    # col1/col2 come from file1['columns'] and file2['columns']
                    relationships.append({
                        'relationship_id': f'rel_{uuid4().hex[:8]}',
                        'file_1': file1['file_name'],
                        'column_1': str(col1),  # original from file
                        'normalized_column_1': str(norm1),
                        'file_2': file2['file_name'],
                        'column_2': str(col2),  # original from file
                        'normalized_column_2': str(norm2),
                        'confidence': round(confidence, 3),
                        'name_similarity': round(name_sim, 3),
                        'value_overlap': round(overlap_ratio, 3),
                        'overlap_count': overlap,
                        'relationship_type': rel_type,
                        'pk_side': pk_side,
                        'pk_column': pk_col,
                        'fk_side': fk_side,
                        'fk_column': fk_col,
                        'sample_overlapping_values': list(vals1 & vals2)[:5]
                    })

    # Sort by confidence descending
    relationships.sort(key=lambda x: x['confidence'], reverse=True)

    # Remove lower confidence duplicates
    # Keep only the best match per column pair
    seen_columns = set()
    filtered = []
    for rel in relationships:
        key1 = f"{rel['file_1']}.{rel['column_1']}"
        key2 = f"{rel['file_2']}.{rel['column_2']}"
        pair = tuple(sorted([key1, key2]))

        if pair not in seen_columns:
            seen_columns.add(pair)
            filtered.append(rel)

    # Detect primary keys
    primary_keys = detect_primary_keys(files)

    return {
        'workspace_id': workspace_id,
        'relationships': filtered,
        'primary_keys': primary_keys,
        'total_relationships': len(filtered),
        'total_files_analyzed': len(files)
    }


def detect_primary_keys(files):
    """
    Detects likely primary key columns
    in each file based on uniqueness.
    Returns ORIGINAL column names only
    (never normalized_column names).
    """
    primary_keys = []

    for file_data in files:
        # Use ORIGINAL columns only, never normalized_columns
        columns = file_data.get('columns', [])
        sample_rows = file_data.get('sample_rows', [])

        if not sample_rows:
            continue

        for col in columns:
            if len(str(col).strip()) <= 1:
                continue

            # Get all values in this column
            values = []
            for row in sample_rows:
                v = row.get(col)
                if v is not None and str(v).strip():
                    values.append(str(v).strip())

            if len(values) < 2:
                continue

            # Check uniqueness
            unique_count = len(set(values))
            total_count = len(values)
            uniqueness_ratio = unique_count / total_count

            # Check for ID-like naming patterns
            col_lower = str(col).lower()
            id_patterns = [
                'id', 'key', 'code', 'num',
                'number', 'no', 'ref', 'uuid',
                'pk', 'sku', 'account'
            ]
            has_id_pattern = any(p in col_lower for p in id_patterns)

            # Score the primary key likelihood
            pk_score = uniqueness_ratio * 0.7
            if has_id_pattern:
                pk_score += 0.3

            if pk_score >= 0.7:
                primary_keys.append({
                    'file': file_data['file_name'],
                    'column': str(col),
                    'uniqueness_ratio': round(uniqueness_ratio, 3),
                    'pk_score': round(pk_score, 3),
                    'confidence': 'high' if pk_score >= 0.9 else 'medium',
                    'reason': (
                        f"{int(uniqueness_ratio*100)}% unique values"
                        + (" + ID naming pattern" if has_id_pattern else "")
                    )
                })

    # Sort by pk_score descending
    primary_keys.sort(key=lambda x: x['pk_score'], reverse=True)

    # Keep only top candidate per file
    seen_files = set()
    top_pks = []
    for pk in primary_keys:
        if pk['file'] not in seen_files:
            seen_files.add(pk['file'])
            top_pks.append(pk)

    return top_pks


def save_relationships_to_db(workspace_id, relationships, db_path):
    """
    Saves detected relationships to SQLite
    for use in HITL review.
    """
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # Clear existing unconfirmed relationships
        # for this workspace
        cursor.execute("""
            DELETE FROM cross_file_relationships
            WHERE workspace_id = ?
            AND confirmed = 0
            AND rejected = 0
        """, (workspace_id,))

        # Insert new relationships
        for rel in relationships:
            cursor.execute("""
                INSERT INTO cross_file_relationships
                (workspace_id, file_1, column_1,
                 file_2, column_2, confidence,
                 relationship_type, confirmed,
                 rejected)
                VALUES (?, ?, ?, ?, ?, ?, ?, 0, 0)
            """, (
                workspace_id,
                rel['file_1'],
                rel['column_1'],
                rel['file_2'],
                rel['column_2'],
                rel['confidence'],
                rel['relationship_type']
            ))

        conn.commit()
        conn.close()
        return True

    except Exception as e:
        return False
