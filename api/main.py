import os
from pathlib import Path

from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from api.routes import upload, chat, quality
from api.routes.workspace import router as workspace_router
from api.routes.commission import router as commission_router

# Project root (parent of api/)
ROOT = Path(__file__).resolve().parent.parent
FRONTEND_DIR = ROOT / "frontend"


def create_app():
    from fastapi import FastAPI
    app = FastAPI(title="Excel LLM Platform")

    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    # Serve frontend files
    frontend_dir = os.path.join(
        os.path.dirname(__file__),
        '..', 'frontend'
    )

    app.mount(
        "/static",
        StaticFiles(directory=frontend_dir),
        name="static"
    )

    @app.get("/")
    def serve_index():
        return FileResponse(
            os.path.join(frontend_dir, 'index.html')
        )

    @app.get("/workspace")
    def serve_workspace():
        return FileResponse(
            os.path.join(frontend_dir, 'workspace.html')
        )

    @app.get("/commission")
    async def commission_page():
        return FileResponse(
            os.path.join(frontend_dir, 'commission.html')
        )

    app.include_router(upload.router)
    app.include_router(chat.router)
    app.include_router(quality.router)
    app.include_router(workspace_router)
    app.include_router(commission_router)

    return app


app = create_app()
