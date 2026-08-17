from contextlib import asynccontextmanager
import logging
from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from .config import settings
from .database import create_db_and_tables, ensure_default_admin, engine as db_engine
from .api import auth, senders, templates, users, recipients, dashboard, upload, tracking, tasks
from .websocket.manager import ws_manager
from .websocket.events import start_dispatcher, stop_dispatcher
from .models.task import Task
from sqlmodel import Session, select
from sqlalchemy import text
from .utils.security import decode_token

logger = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(app: FastAPI):
    engine = None
    try:
        create_db_and_tables()
        ensure_default_admin()

        # Start websocket event bridge
        await start_dispatcher()

        # Auto-recover stuck tasks on startup
        from .services.send_engine import get_send_engine
        engine = get_send_engine()
        with Session(db_engine) as session:
            stuck_tasks = session.exec(
                select(Task)
                .where(Task.status.in_(["running", "paused"]))  # type: ignore[attr-defined]
                .order_by(Task.id)
                .limit(1000)
            ).all()
            for task in stuck_tasks:
                if task.status == "paused" and task.pause_reason == "manual":
                    continue
                # only recover if pending logs remain
                from .models.send_log import SendLog
                pending = session.exec(
                    select(SendLog)
                    .where(SendLog.task_id == task.id, SendLog.status == "pending")
                    .limit(1)
                ).first()
                if pending:
                    engine.submit(task.id)
                    logger.info("Auto-recovered task %s on startup", task.id)
        yield
    finally:
        # Always release background workers and connection pools, including
        # when the ASGI server is cancelled or lifespan exits with an error.
        if engine is not None:
            try:
                engine.shutdown()
            except Exception:
                logger.exception("Send engine shutdown failed")
        try:
            await stop_dispatcher()
        except Exception:
            logger.exception("WebSocket dispatcher shutdown failed")
        try:
            db_engine.dispose()
        except Exception:
            logger.exception("Database engine shutdown failed")


app = FastAPI(
    title=settings.APP_NAME,
    description="企业级邮箱群发系统 API",
    version="4.0.0",
    lifespan=lifespan,
    docs_url="/docs" if settings.ENABLE_DOCS else None,
    redoc_url="/redoc" if settings.ENABLE_DOCS else None,
    openapi_url="/openapi.json" if settings.ENABLE_DOCS else None,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins_list,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Register API routers
app.include_router(auth.router)
app.include_router(senders.router)
app.include_router(templates.router)
app.include_router(users.router)
app.include_router(recipients.router)
app.include_router(dashboard.router)
app.include_router(upload.router)
app.include_router(tracking.router)
app.include_router(tasks.router)


@app.get("/health")
def health_check():
    return {"status": "ok"}


@app.get("/health/ready")
def readiness_check():
    redis_required = str(settings.RATE_LIMIT_BACKEND or "auto").lower() == "redis"
    checks = {"database": False}
    try:
        with db_engine.connect() as connection:
            connection.execute(text("SELECT 1"))
        checks["database"] = True
    except Exception:
        logger.warning("Database readiness check failed", exc_info=True)
    if redis_required:
        checks["redis"] = False
        try:
            import redis

            redis.Redis.from_url(
                settings.REDIS_URL,
                socket_connect_timeout=0.5,
                socket_timeout=0.5,
            ).ping()
            checks["redis"] = True
        except Exception:
            logger.warning("Redis readiness check failed", exc_info=True)
    if not all(checks.values()):
        return JSONResponse(status_code=503, content={"status": "not_ready", "checks": checks})
    return {"status": "ready", "checks": checks}


# WebSocket endpoint
@app.websocket("/ws/tasks/{task_id}")
async def websocket_task(websocket: WebSocket, task_id: int):
    token: str | None = None
    selected_protocol = None
    for protocol in websocket.headers.get("sec-websocket-protocol", "").split(","):
        protocol = protocol.strip()
        if protocol.startswith("access-token."):
            token = protocol.removeprefix("access-token.")
            selected_protocol = protocol
            break
    # Keep compatibility with older clients while new clients avoid putting
    # bearer tokens in access-log URLs.
    if not token:
        token = websocket.query_params.get("token", "")
    payload = decode_token(token) if token else None
    if not payload or payload.get("type") != "access" or not payload.get("sub"):
        await websocket.close(code=4401, reason="Authentication required")
        return

    with Session(db_engine) as session:
        task = session.get(Task, task_id)
        if not task:
            await websocket.close(code=4404, reason="Task not found")
            return
        if task.user_id != payload["sub"]:
            await websocket.close(code=4403, reason="Task access denied")
            return

    await ws_manager.connect(websocket, task_id, subprotocol=selected_protocol)
    try:
        while True:
            # Keep connection alive, client can send heartbeats
            await websocket.receive_text()
    except WebSocketDisconnect:
        pass
    finally:
        ws_manager.disconnect(websocket, task_id)
