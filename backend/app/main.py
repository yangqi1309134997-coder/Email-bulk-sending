from contextlib import asynccontextmanager
from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from .config import settings
from .database import create_db_and_tables
from .api import auth, senders, templates, users, recipients, dashboard, upload, tracking, tasks
from .websocket.manager import ws_manager


@asynccontextmanager
async def lifespan(app: FastAPI):
    create_db_and_tables()
    yield


app = FastAPI(
    title=settings.APP_NAME,
    description="企业级邮箱群发系统 API",
    version="4.0.0",
    lifespan=lifespan,
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


# WebSocket endpoint
@app.websocket("/ws/tasks/{task_id}")
async def websocket_task(websocket: WebSocket, task_id: int):
    await ws_manager.connect(websocket, task_id)
    try:
        while True:
            # Keep connection alive, client can send heartbeats
            data = await websocket.receive_text()
            # Optional: handle client messages
    except WebSocketDisconnect:
        ws_manager.disconnect(websocket, task_id)