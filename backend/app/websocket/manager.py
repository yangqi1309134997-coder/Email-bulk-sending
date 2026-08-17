from fastapi import WebSocket
from typing import Dict, Set
import json


class ConnectionManager:
    def __init__(self):
        self.active_connections: Dict[int, Set[WebSocket]] = {}

    async def connect(self, websocket: WebSocket, task_id: int, subprotocol: str | None = None):
        await websocket.accept(subprotocol=subprotocol)
        if task_id not in self.active_connections:
            self.active_connections[task_id] = set()
        self.active_connections[task_id].add(websocket)

    def disconnect(self, websocket: WebSocket, task_id: int):
        if task_id in self.active_connections:
            self.active_connections[task_id].discard(websocket)
            if not self.active_connections[task_id]:
                del self.active_connections[task_id]

    async def send_to_task(self, task_id: int, message: dict):
        if task_id in self.active_connections:
            data = json.dumps(message, ensure_ascii=False, default=str)
            dead = set()
            for ws in self.active_connections[task_id]:
                try:
                    await ws.send_text(data)
                except Exception:
                    dead.add(ws)
            self.active_connections[task_id] -= dead
            if not self.active_connections[task_id]:
                self.active_connections.pop(task_id, None)

    async def broadcast(self, message: dict):
        data = json.dumps(message, ensure_ascii=False, default=str)
        for task_id in list(self.active_connections.keys()):
            dead = set()
            for ws in self.active_connections[task_id]:
                try:
                    await ws.send_text(data)
                except Exception:
                    dead.add(ws)
            self.active_connections[task_id] -= dead
            if not self.active_connections[task_id]:
                self.active_connections.pop(task_id, None)


ws_manager = ConnectionManager()
