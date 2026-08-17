"""Compatibility Celery task that delegates to the persistent SendEngine.

Historical design used Celery workers for batch sending. The system now uses an
in-process SendEngine for better connection pooling and risk-control state.
This task remains for Celery beat / external triggers.
"""

import time
from time import monotonic

from celery import shared_task
from sqlmodel import Session

from ..config import settings
from ..database import engine
from ..models.task import Task
from ..services.send_engine import get_send_engine


@shared_task(
    bind=True,
    max_retries=3,
    acks_late=True,
    reject_on_worker_lost=True,
    name="app.tasks.send_email.send_batch_task",
)
def send_batch_task(self, task_id: int):
    task_id = int(task_id)
    get_send_engine().submit(task_id)
    deadline = monotonic() + max(
        60.0, float(getattr(settings, "CELERY_TASK_WAIT_SECONDS", 86400))
    )
    while monotonic() < deadline:
        with Session(engine) as session:
            task = session.get(Task, task_id)
            if task is None:
                return {"task_id": task_id, "status": "missing"}
            current_status = task.status
            pause_reason = getattr(task, "pause_reason", "")
        if current_status in {"completed", "failed", "cancelled"}:
            return {"task_id": task_id, "status": current_status}
        if current_status == "paused" and pause_reason == "manual":
            return {"task_id": task_id, "status": current_status}
        time.sleep(1.0)

    # Leave the durable task state intact. Late acknowledgement means the
    # broker can redeliver this monitoring task after a worker restart.
    raise RuntimeError(f"Timed out waiting for send task {task_id}")
