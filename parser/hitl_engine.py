import sqlite3
import json
import re
import statistics
from uuid import uuid4
from difflib import SequenceMatcher


def detect_review_items(workbook_id, db_path):

    review_items = []

    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # Load workbook parsed data
        cursor.execute("""
            SELECT parsed_json
            FROM workbooks
            WHERE workbook_id = ?
        """, (workbook_id,))
        row = cursor.fetchone()

        if not row:
            return []

        parsed = json.loads(row['parsed_json'] or '{}')
        sheets_list = parsed.get('sheets', [])

        # Build sheets dict from parsed_json structure
        sheets = {}
        for sheet in sheets_list:
            sheet_name = sheet.get('sheet_name', '')
            if not sheet_name:
                continue
            sheets[sheet_name] = {
                'columns': sheet.get('columns', []),
                'sample_rows': sheet.get('sample_rows', [])
            }

        # Load schema mappings separately
        cursor.execute("""
            SELECT sheet_name, entity, confidence,
                   needs_review, key_metrics
            FROM schema_mappings
            WHERE workbook_id = ?
        """, (workbook_id,))
        mappings = {}
        for row in cursor.fetchall():
            mappings[row['sheet_name']] = dict(row)

        conn.close()

    except Exception as e:
        return []

    # ─────────────────────────────────────
    # DETECTOR 1: Same metric across sheets
    # ─────────────────────────────────────
    try:
        metric_map = {}

        for sheet_name, sheet_data in sheets.items():
            for row in sheet_data['sample_rows']:
                # Find the label value in this row
                # Check column A first then any
                # string value in first position
                label = None
                for col in sheet_data['columns']:
                    val = row.get(col, '')
                    if (val and
                        isinstance(val, str) and
                        len(val.strip()) > 2 and
                        len(col.strip()) <= 2):
                        label = val.strip()
                        break

                if not label:
                    continue

                label_key = label.lower().strip()

                if label_key not in metric_map:
                    metric_map[label_key] = []

                # Get numeric values from this row
                numeric_vals = {}
                for col, val in row.items():
                    if col != col[:2]:
                        continue
                    try:
                        numeric_vals[col] = float(val)
                    except (TypeError, ValueError):
                        pass

                metric_map[label_key].append({
                    'sheet': sheet_name,
                    'label': label,
                    'values': numeric_vals
                })

        # Flag metrics appearing in 2+ sheets
        for label_key, appearances in metric_map.items():
            if len(appearances) < 2:
                continue

            # Check if values match across sheets
            all_values = []
            for a in appearances:
                all_values.extend(
                    a['values'].values()
                )

            values_match = True
            if len(all_values) >= 2:
                max_v = max(all_values)
                min_v = min(all_values)
                if max_v != 0:
                    diff_pct = abs(max_v - min_v) / abs(max_v)
                    values_match = diff_pct < 0.01

            review_items.append({
                'review_id': f'rev_{uuid4().hex[:8]}',
                'type': 'same_metric_across_sheets',
                'priority': 'high' if not values_match else 'info',
                'question': f'Are these all the same metric?',
                'finding': {
                    'metric_name': appearances[0]['label'],
                    'appearances': [
                        {'sheet': a['sheet']}
                        for a in appearances
                    ],
                    'values_match': values_match
                },
                'options': [
                    'Yes same metric - use highest confidence sheet as primary',
                    'No they are different metrics',
                    'They are related but different'
                ],
                'impact': 'Helps system always find the right source for this metric'
            })

    except Exception:
        pass

    # ─────────────────────────────────────
    # DETECTOR 2: Column identity
    # ─────────────────────────────────────
    try:
        sheet_names = list(sheets.keys())

        for i in range(len(sheet_names)):
            for j in range(i + 1, len(sheet_names)):
                sheet1 = sheet_names[i]
                sheet2 = sheet_names[j]

                cols1 = [
                    c for c in sheets[sheet1]['columns']
                    if len(str(c).strip()) > 1
                    and str(c).strip().lower() != 'label'
                ]
                cols2 = [
                    c for c in sheets[sheet2]['columns']
                    if len(str(c).strip()) > 1
                    and str(c).strip().lower() != 'label'
                ]

                for col1 in cols1:
                    for col2 in cols2:
                        if str(col1).lower() == str(col2).lower():
                            continue

                        name_sim = SequenceMatcher(
                            None,
                            str(col1).lower().strip(),
                            str(col2).lower().strip()
                        ).ratio()

                        # Get unique values from each column
                        vals1 = set()
                        for row in sheets[sheet1]['sample_rows']:
                            v = row.get(col1, '')
                            if v and str(v).strip():
                                vals1.add(str(v).strip().lower())

                        vals2 = set()
                        for row in sheets[sheet2]['sample_rows']:
                            v = row.get(col2, '')
                            if v and str(v).strip():
                                vals2.add(str(v).strip().lower())

                        if not vals1 or not vals2:
                            continue

                        overlap = len(vals1 & vals2)
                        overlap_ratio = overlap / min(
                            len(vals1), len(vals2)
                        )

                        combined = (name_sim * 0.4 +
                                    overlap_ratio * 0.6)

                        if 0.25 < combined < 0.95:
                            review_items.append({
                                'review_id': f'rev_{uuid4().hex[:8]}',
                                'type': 'column_identity',
                                'priority': 'medium',
                                'question': f'Are these the same field?',
                                'finding': {
                                    'sheet_1': sheet1,
                                    'column_1': str(col1),
                                    'sample_values_1': list(vals1)[:5],
                                    'sheet_2': sheet2,
                                    'column_2': str(col2),
                                    'sample_values_2': list(vals2)[:5],
                                    'overlap_count': overlap,
                                    'ai_confidence': round(combined, 2)
                                },
                                'options': [
                                    'Yes they are the same field',
                                    'No they are different',
                                    'They are related but different'
                                ],
                                'impact': 'Enables cross-sheet questions joining these two sheets'
                            })

    except Exception:
        pass

    # ─────────────────────────────────────
    # DETECTOR 3: Date format ambiguity
    # ─────────────────────────────────────
    try:
        ambiguous_pattern = re.compile(
            r'^\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}$'
        )

        for sheet_name, sheet_data in sheets.items():
            for col in sheet_data['columns']:
                sample_vals = []
                for row in sheet_data['sample_rows']:
                    v = str(row.get(col, '')).strip()
                    if v and ambiguous_pattern.match(v):
                        sample_vals.append(v)
                    if len(sample_vals) >= 5:
                        break

                if len(sample_vals) >= 3:
                    review_items.append({
                        'review_id': f'rev_{uuid4().hex[:8]}',
                        'type': 'date_format',
                        'priority': 'medium',
                        'question': f'How should we read these dates?',
                        'finding': {
                            'sheet': sheet_name,
                            'column': str(col),
                            'sample_values': sample_vals
                        },
                        'options': [
                            'MM/DD/YYYY',
                            'DD/MM/YYYY',
                            'YYYY/MM/DD'
                        ],
                        'impact': 'Ensures dates are interpreted correctly in all analysis'
                    })

    except Exception:
        pass

    # ─────────────────────────────────────
    # DETECTOR 4: Unit scale ambiguity
    # ─────────────────────────────────────
    try:
        for sheet_name, sheet_data in sheets.items():
            for col in sheet_data['columns']:
                numeric_values = []
                for row in sheet_data['sample_rows']:
                    v = row.get(col)
                    if v is None or v == '':
                        continue
                    try:
                        fv = float(v)
                        if fv > 0:
                            numeric_values.append(fv)
                    except (TypeError, ValueError):
                        pass

                if len(numeric_values) < 3:
                    continue

                median_val = statistics.median(numeric_values)
                max_val = max(numeric_values)
                min_val = min(numeric_values)

                if not (0 < median_val < 100000):
                    continue

                if max_val / median_val >= 1000:
                    continue

                # Skip if clearly small counts
                all_integers = all(
                    v == int(v) for v in numeric_values
                )
                small_range = max_val <= 1000

                if all_integers and small_range:
                    continue

                review_items.append({
                    'review_id': f'rev_{uuid4().hex[:8]}',
                    'type': 'unit_scale',
                    'priority': 'medium',
                    'question': f'What unit are the numbers in column "{col}"?',
                    'finding': {
                        'sheet': sheet_name,
                        'column': str(col),
                        'sample_values': numeric_values[:5],
                        'median_value': median_val
                    },
                    'options': [
                        'Exact value (e.g. $500)',
                        'Thousands (e.g. $500K)',
                        'Millions (e.g. $0.5M)'
                    ],
                    'impact': 'Ensures all calculations use the correct scale'
                })

    except Exception:
        pass

    # ─────────────────────────────────────
    # DETECTOR 5: Low confidence sheets
    # ─────────────────────────────────────
    try:
        for sheet_name, mapping in mappings.items():
            confidence = mapping.get('confidence', 1.0)
            needs_review = mapping.get('needs_review', False)

            if confidence < 0.75 or needs_review:
                entity = mapping.get('entity', 'unknown')
                cols = json.loads(
                    mapping.get('key_metrics', '[]') or '[]'
                )

                review_items.append({
                    'review_id': f'rev_{uuid4().hex[:8]}',
                    'type': 'low_confidence_sheet',
                    'priority': 'low',
                    'question': f'We think sheet "{sheet_name}" contains {entity} data. Is that correct?',
                    'finding': {
                        'sheet': sheet_name,
                        'ai_guess': entity,
                        'confidence': confidence,
                        'sample_columns': cols[:5] if isinstance(cols, list) else []
                    },
                    'options': [
                        f'Yes it contains {entity}',
                        'No it contains something else',
                        'Skip this sheet in analysis'
                    ],
                    'impact': 'Improves accuracy of sheet identification'
                })

    except Exception:
        pass

    # Sort by priority and cap at 10
    priority_order = {'high': 0, 'info': 1,
                      'medium': 2, 'low': 3}
    review_items.sort(
        key=lambda x: priority_order.get(
            x['priority'], 99
        )
    )

    return review_items[:10]
