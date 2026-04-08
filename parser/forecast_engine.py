"""Generic time series forecasting via linear regression. Works on any spreadsheet domain."""
import re
from typing import Any

import numpy as np


def _detect_sequential_columns(columns: list[str]) -> list[str]:
    """Detect columns that follow a sequential time/period pattern."""
    sequential_cols: list[str] = []
    for col in columns:
        col_lower = str(col).lower().strip()
        if re.search(r"year\s*\d+", col_lower):
            sequential_cols.append(col)
        elif re.search(r"q\d+", col_lower):
            sequential_cols.append(col)
        elif re.search(r"month\s*\d+", col_lower):
            sequential_cols.append(col)
        elif re.search(r"^\d{4}$", col_lower):
            sequential_cols.append(col)
        elif re.search(r"period\s*\d+", col_lower):
            sequential_cols.append(col)
        elif re.search(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b", col_lower):
            sequential_cols.append(col)
    return sequential_cols


def _projected_names(last_col: str, sequential_cols: list[str], periods: int) -> list[str]:
    """Determine naming convention for projected periods from the last column."""
    last_lower = str(last_col).lower()
    num_match = re.search(r"\d+", last_col)

    if "year" in last_lower:
        if num_match:
            last_num = int(num_match.group())
            return [
                last_col.replace(num_match.group(), str(last_num + i))
                for i in range(1, periods + 1)
            ]
    elif "q" in last_lower and num_match:
        last_num = int(num_match.group())
        return [f"Q{last_num + i}" for i in range(1, periods + 1)]
    elif "month" in last_lower and num_match:
        last_num = int(num_match.group())
        return [
            last_col.replace(num_match.group(), str(last_num + i))
            for i in range(1, periods + 1)
        ]
    elif re.search(r"^\d{4}$", str(last_col)):
        last_num = int(last_col)
        return [str(last_num + i) for i in range(1, periods + 1)]

    return [f"Period {len(sequential_cols) + i}" for i in range(1, periods + 1)]


def run_forecast(
    sheets_data: list[dict[str, Any]],
    periods: int = 2,
) -> list[dict[str, Any]]:
    """
    Run linear regression forecasting on sheets with sequential numeric columns.
    Returns list of forecast results per sheet, or empty list if no forecastable data.
    """
    if not sheets_data:
        return []

    results: list[dict[str, Any]] = []

    try:
        for sheet in sheets_data:
            sheet_name = sheet.get("sheet_name") or sheet.get("name") or "Unknown"
            columns = sheet.get("columns") or sheet.get("clean_columns") or []
            rows = sheet.get("rows") or sheet.get("clean_rows") or sheet.get("sample_rows") or []

            if not columns or not rows:
                continue

            # STEP 1: Detect sequential columns
            sequential_cols = _detect_sequential_columns(columns)
            if len(sequential_cols) < 4:
                continue

            # STEP 2 & 3 & 4 & 5 & 6: Process each row
            forecasts: list[dict[str, Any]] = []
            for row_index, row in enumerate(rows):
                values = []
                for col in sequential_cols:
                    val = row.get(col)
                    if val is not None and isinstance(val, (int, float)):
                        values.append(float(val))
                    else:
                        values.append(None)

                non_null = [v for v in values if v is not None]
                if len(non_null) < 4:
                    continue

                if all(-2 <= v <= 2 for v in non_null):
                    continue

                if len(set(non_null)) == 1:
                    continue

                # STEP 3: Fit linear regression
                x = list(range(1, len(non_null) + 1))
                y = non_null

                slope, intercept = np.polyfit(x, y, 1)

                y_pred = [slope * xi + intercept for xi in x]
                ss_res = sum((yi - yp) ** 2 for yi, yp in zip(y, y_pred))
                ss_tot = sum((yi - np.mean(y)) ** 2 for yi in y)

                if ss_tot == 0:
                    continue

                r_squared = 1 - (ss_res / ss_tot)

                if r_squared < 0.85:
                    continue

                # STEP 4: Project next periods
                last_col = sequential_cols[-1]
                projected_names_list = _projected_names(last_col, sequential_cols, periods)

                projected: dict[str, float] = {}
                for i, name in enumerate(projected_names_list):
                    next_x = len(non_null) + i + 1
                    projected_val = slope * next_x + intercept
                    projected[name] = round(float(projected_val), 2)

                # STEP 5: Row label
                row_label = row.get("label") or row.get("name") or row.get("category") or f"Row {row_index}"

                # STEP 6: Build result
                historical = {
                    col: row.get(col)
                    for col in sequential_cols
                    if row.get(col) is not None
                }

                forecast_item = {
                    "row_label": str(row_label),
                    "historical": historical,
                    "projected": projected,
                    "slope": round(float(slope), 2),
                    "r_squared": round(float(r_squared), 4),
                    "confidence": (
                        "high" if r_squared >= 0.95 else "medium" if r_squared >= 0.85 else "low"
                    ),
                    "trend_direction": "increasing" if slope > 0 else "decreasing",
                    "avg_change_per_period": round(float(slope), 2),
                }
                forecasts.append(forecast_item)

            if not forecasts:
                continue

            # STEP 7: Top 5 by r_squared
            forecasts_sorted = sorted(forecasts, key=lambda f: f["r_squared"], reverse=True)
            top_5 = forecasts_sorted[:5]

            results.append({
                "sheet": sheet_name,
                "sequential_columns": sequential_cols,
                "projected_periods": periods,
                "forecasts": top_5,
            })

    except Exception:
        return []

    return results
