from fastapi import APIRouter, HTTPException

router = APIRouter(prefix="/quality", tags=["quality"])


@router.get("/{workbook_id}")
async def get_quality(workbook_id: str):
    """Return quality metrics for a parsed workbook."""
    # TODO: use quality_engine for workbook_id
    return {"workbook_id": workbook_id, "metrics": {}}
