import logging
import traceback

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.database import Base, engine
from app.routers import admin, auth, check_submission, dashboard, executions, forms_generator, init_project, tools
from app.services.storage_service import ensure_dirs

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Pure Tools Platform",
    description="Platform to manage and run Pure parametric tools",
    version="1.0.0",
)


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    logger.error(f"Unhandled error: {exc}\n{traceback.format_exc()}")
    return JSONResponse(status_code=500, content={"detail": str(exc)})

# CORS for React frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://localhost:5174",
        "http://localhost:5175",
        "http://localhost:5176",
        "http://localhost:5177",
        "http://localhost:5178",
        "http://127.0.0.1:5173",
        "http://127.0.0.1:5174",
        "http://127.0.0.1:5175",
        "http://127.0.0.1:5176",
        "http://127.0.0.1:5177",
        "http://127.0.0.1:5178",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)

# Routers
app.include_router(auth.router)
app.include_router(tools.router)
app.include_router(executions.router)
app.include_router(init_project.router)
app.include_router(forms_generator.router)
app.include_router(check_submission.router)
app.include_router(admin.router)
app.include_router(dashboard.router)


@app.on_event("startup")
def on_startup():
    try:
        # Import all models so they are registered with Base
        from app.models import tool_execution  # noqa: F401
        Base.metadata.create_all(bind=engine)
        ensure_dirs()
        logger.info("Database tables created successfully")
    except Exception as e:
        logger.error(f"Failed to connect to database: {e}")


@app.get("/api/health")
def health_check():
    return {"status": "ok"}
