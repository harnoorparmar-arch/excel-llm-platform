from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse, StreamingResponse
import os
import shutil
import uuid
import time
import gc
from datetime import datetime
from pathlib import Path

router = APIRouter(prefix="/commission", tags=["commission"])

# Temp upload directory
UPLOAD_DIR = Path("uploads/commission")
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

DB_PATH = "storage/workbooks.db"

ALLOWED_COMMISSION_EXTENSIONS = {
    ".pdf",
    ".xlsx",
    ".xls",
    ".xlsm",
    ".xlsb",
    ".csv",
    ".tsv",
    ".txt",
    ".slk",
}


def safe_remove(path, retries=5, delay=0.5):
    for attempt in range(retries):
        try:
            os.remove(path)
            return
        except PermissionError:
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                print(f"Warning: Could not delete temp file: {path}")


@router.post("/upload")
async def upload_commission_file(
    file: UploadFile = File(...)
):
    """
    Upload and process a commission file.
    Returns extracted PO groups with
    confidence scores ready for HITL review.
    """
    file_ext = Path(file.filename).suffix.lower()
    if file_ext not in ALLOWED_COMMISSION_EXTENSIONS:
        raise HTTPException(
            status_code=400,
            detail=(
                "Unsupported file type. Accepted: "
                ".pdf .xlsx .xls .xlsm .xlsb .csv .tsv .txt .slk"
            ),
        )

    # Save uploaded file
    job_id = str(uuid.uuid4())[:8]
    save_path = UPLOAD_DIR / f"{job_id}{file_ext}"

    try:
        with open(save_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Failed to save file: {e}"
        )

    # Process the file
    try:
        from parser.commission_extractor import (
            process_commission_file
        )

        result = process_commission_file(
            str(save_path),
            db_path=DB_PATH,
            original_filename=file.filename
        )

        if result['status'] == 'error':
            raise HTTPException(
                status_code=500,
                detail=result['error']
            )

        if result['status'] == 'duplicate':
            return JSONResponse({
                "status": "duplicate",
                "message": "This file has already been processed",
                "job_id": job_id
            })

        # Format response for UI
        grouped = result['grouped']

        po_groups = []
        for po_norm, group in grouped.items():
            po_groups.append({
                "po_number": group['po_number'],
                "po_normalized": po_norm,
                "dealer_name": group['dealer_name'],
                "total_commission": group['total_commission'],
                "total_sale_amount": group['total_sale_amount'],
                "line_item_count": group['line_item_count'],
                "invoice_numbers": group['invoice_numbers'],
                "invoice_dates": group['invoice_dates'],
                "needs_review": group['needs_review'],
                "min_confidence": group['min_confidence'],
                "all_issues": list(set(group['all_issues'])),
                "line_items": group['line_items'],
                "commission_origination": group.get('commission_origination'),
                "commission_specification": group.get('commission_specification'),
                "commission_destination": group.get('commission_destination'),
                "has_commission_splits": group.get('has_commission_splits', False),
                "commission_rebate": group.get('commission_rebate'),
            })

        # Sort: review items first
        po_groups.sort(
            key=lambda x: (
                not x['needs_review'],
                -x['total_commission']
            )
        )

        return JSONResponse({
            "status": "success",
            "job_id": job_id,
            "document_hash": result['document']['hash'],
            "file_name": file.filename,
            "summary": result['summary'],
            "po_groups": po_groups,
            "no_po_items": result['no_po_items'],
            "processed_at": datetime.now().isoformat()
        })

    except HTTPException:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(
            status_code=500,
            detail=str(e)
        )
    finally:
        # Clean up temp file
        if save_path.exists():
            gc.collect()
            time.sleep(0.1)
            safe_remove(save_path)


@router.post("/export-csv")
async def export_to_csv(request: dict):
    """
    Export approved PO groups to CSV.
    One row per invoice (not per PO).
    """
    import csv
    import io
    from collections import defaultdict

    manufacturer = request.get('manufacturer', '')
    period = request.get('period', '')
    approved_pos = request.get('approved_pos', [])

    output = io.StringIO()
    writer = csv.writer(output)

    writer.writerow([
        'PO Number', 'Dealer Name', 'Invoice Number', 'Invoice Date',
        'Sale Amount', 'Commission Amount', 'Origination', 'Specification',
        'Destination', 'Manufacturer', 'Period', 'Source File', 'PO Source'
    ])

    for po in approved_pos:
        line_items = po.get('line_items', [])

        if line_items:
            invoice_groups = defaultdict(list)
            for item in line_items:
                inv = str(item.get('invoice_number') or 'NO_INVOICE')
                invoice_groups[inv].append(item)

            for inv_num, inv_items in invoice_groups.items():
                sale_total = sum(
                    item.get('sale_amount') or 0
                    for item in inv_items
                )
                comm_total = sum(
                    item.get('commission_amount') or 0
                    for item in inv_items
                )
                orig_total = sum(
                    item.get('commission_origination') or 0
                    for item in inv_items
                )
                spec_total = sum(
                    item.get('commission_specification') or 0
                    for item in inv_items
                )
                dest_total = sum(
                    item.get('commission_destination') or 0
                    for item in inv_items
                )
                inv_date = inv_items[0].get('invoice_date', '')
                inv_display = inv_num if inv_num != 'NO_INVOICE' else ''

                writer.writerow([
                    po.get('po_number', ''),
                    po.get('dealer_name', ''),
                    inv_display,
                    inv_date,
                    round(sale_total, 2),
                    round(comm_total, 2),
                    round(orig_total, 2) if orig_total else '',
                    round(spec_total, 2) if spec_total else '',
                    round(dest_total, 2) if dest_total else '',
                    manufacturer,
                    period,
                    po.get('source_file', ''),
                    po.get('po_source', 'extracted')
                ])
        else:
            writer.writerow([
                po.get('po_number', ''),
                po.get('dealer_name', ''),
                ', '.join(str(i) for i in po.get('invoice_numbers', []) if i),
                ', '.join(str(d) for d in po.get('invoice_dates', []) if d),
                po.get('total_sale_amount', ''),
                po.get('total_commission', ''),
                po.get('commission_origination', ''),
                po.get('commission_specification', ''),
                po.get('commission_destination', ''),
                manufacturer,
                period,
                po.get('source_file', ''),
                po.get('po_source', 'extracted')
            ])

    ra_list = request.get("rebates_adjustments") or []
    if ra_list:
        writer.writerow([])
        writer.writerow([
            "TYPE",
            "Invoice Name",
            "Dealer",
            "Factory",
            "Sales Person",
            "Invoice Date",
            "Invoice Amount",
            "Commissionable Sales",
            "Origination",
            "Specification",
            "Destination",
            "Paid Commission",
            "Check #",
            "Payment Date",
            "Paid Status",
        ])
        for ra in ra_list:
            writer.writerow([
                str(ra.get("type", "")).upper(),
                ra.get("invoice_name", ""),
                ra.get("dealer", ""),
                ra.get("factory", ""),
                ra.get("sales_person", ""),
                ra.get("invoice_date", ""),
                ra.get("invoice_amount", ""),
                ra.get("commissionable_sales", ""),
                ra.get("origination_amount", ""),
                ra.get("spec_amount", ""),
                ra.get("dest_amount", ""),
                ra.get("paid_commission", ""),
                ra.get("check_number", ""),
                ra.get("payment_date", ""),
                ra.get("paid_status", ""),
            ])

    output.seek(0)
    safe_mfr = manufacturer.replace(' ', '_').replace('/', '_')
    safe_period = period.replace(' ', '_')
    filename = f"commission_{safe_mfr}_{safe_period}.csv"

    return StreamingResponse(
        iter([output.getvalue()]),
        media_type='text/csv',
        headers={
            'Content-Disposition': f'attachment; filename="{filename}"'
        }
    )


@router.post("/file")
async def file_to_salesforce(request: dict):
    """
    File approved commission POs to Salesforce.
    Marks document as processed on success for duplicate detection.
    """
    # Actual Salesforce filing logic to be implemented
    # For now, acknowledge and mark as processed
    try:
        from parser.commission_extractor import mark_as_processed

        if request.get('document_hash'):
            mark_as_processed(
                request['document_hash'],
                request.get('file_name', ''),
                DB_PATH
            )

        return JSONResponse({
            "status": "success",
            "message": "Filed successfully",
            "approved_count": len(request.get('approved_pos', []))
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


@router.get("/templates")
async def list_templates():
    """List all saved manufacturer templates."""
    import sqlite3
    import json

    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT manufacturer, mapping,
                   created_at, updated_at,
                   used_count
            FROM commission_templates
            ORDER BY manufacturer
        """)
        rows = cursor.fetchall()
        conn.close()

        templates = []
        for row in rows:
            mapping = json.loads(row[1])
            templates.append({
                "manufacturer": row[0],
                "columns_mapped": {
                    k: v for k, v in mapping.items()
                    if k not in [
                        'skip_rows_where',
                        'primary_sheet',
                        'skip_sheets',
                        'manufacturer',
                        'period'
                    ]
                },
                "primary_sheet": mapping.get(
                    'primary_sheet'
                ),
                "created_at": row[2],
                "updated_at": row[3],
                "used_count": row[4]
            })

        return JSONResponse({"templates": templates})

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


@router.delete("/templates/{manufacturer}")
async def delete_template(manufacturer: str):
    """
    Delete a saved template.
    Forces re-mapping on next file upload.
    """
    import sqlite3

    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            "DELETE FROM commission_templates "
            "WHERE manufacturer = ?",
            (manufacturer.lower(),)
        )
        deleted = cursor.rowcount
        conn.commit()
        conn.close()

        return JSONResponse({
            "deleted": deleted > 0,
            "manufacturer": manufacturer
        })

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=str(e)
        )
