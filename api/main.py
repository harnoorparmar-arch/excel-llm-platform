import os
from pathlib import Path

from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from api.routes.commission import router as commission_router

ROOT = Path(__file__).resolve().parent.parent
FRONTEND_DIR = ROOT / "frontend"


def create_app():
    from fastapi import FastAPI

    app = FastAPI(title="Commission Filing")

    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    frontend_dir = os.path.join(os.path.dirname(__file__), "..", "frontend")

    app.mount(
        "/static",
        StaticFiles(directory=frontend_dir),
        name="static",
    )

    @app.get("/")
    def serve_root():
        return RedirectResponse(url="/commission")

    @app.get("/commission")
    async def commission_page():
        return FileResponse(os.path.join(frontend_dir, "commission.html"))

    app.include_router(commission_router)

    return app


app = create_app()
