"""Thread-safe progress event bus bridging sync SendEngine to async WebSocket."""

from __future__ import annotations

import asyncio
import logging
import queue
import threading
from typing import Any, Optional

logger = logging.getLogger(__name__)

# A bounded queue prevents a burst of progress events from exhausting process
# memory. The sender state remains durable in SQL, so dropping an event only
# affects live UI updates and the client can refresh the task on reconnect.
_event_queue: queue.Queue[tuple[int, dict[str, Any]]] = queue.Queue(maxsize=10000)
_loop: Optional[asyncio.AbstractEventLoop] = None
_dispatcher_task: Optional[asyncio.Task] = None
_lock = threading.Lock()


def set_event_loop(loop: asyncio.AbstractEventLoop) -> None:
    global _loop
    with _lock:
        _loop = loop


def publish_task_event(task_id: int, message: dict[str, Any]) -> None:
    """Publish a progress event from any thread."""
    payload = dict(message or {})
    payload.setdefault("task_id", task_id)
    payload.setdefault("type", "progress")
    try:
        _event_queue.put_nowait((int(task_id), payload))
    except queue.Full:
        # Live progress is best-effort; durable task state remains in SQL.
        logger.debug("WebSocket progress queue full; dropping event for %s", task_id)
    except Exception:
        logger.exception("Failed to queue task event for %s", task_id)


async def _dispatcher_loop() -> None:
    from .manager import ws_manager

    while True:
        drained = 0
        latest_by_task: dict[int, tuple[int, dict[str, Any]]] = {}
        while drained < 1000:
            try:
                task_id, message = _event_queue.get_nowait()
            except queue.Empty:
                break
            latest_by_task[task_id] = (drained, message)
            drained += 1
        for task_id, (_, message) in sorted(
            latest_by_task.items(), key=lambda item: item[1][0]
        ):
            try:
                await ws_manager.send_to_task(task_id, message)
            except Exception:
                logger.exception("Failed to push websocket event for task %s", task_id)
        await asyncio.sleep(0.1)


async def start_dispatcher() -> None:
    global _dispatcher_task
    loop = asyncio.get_running_loop()
    set_event_loop(loop)
    if _dispatcher_task is None or _dispatcher_task.done():
        _dispatcher_task = asyncio.create_task(_dispatcher_loop())
        logger.info("WebSocket event dispatcher started")


async def stop_dispatcher() -> None:
    global _dispatcher_task
    if _dispatcher_task and not _dispatcher_task.done():
        _dispatcher_task.cancel()
        try:
            await _dispatcher_task
        except (asyncio.CancelledError, Exception):
            pass
    _dispatcher_task = None
