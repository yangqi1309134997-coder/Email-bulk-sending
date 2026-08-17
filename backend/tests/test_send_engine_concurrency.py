"""Tests for concurrent send engine stability and risk control."""

import json
import time
from datetime import timedelta

from sqlmodel import Session, select

from app.database import engine, create_db_and_tables
from app.models.sender import Sender
from app.models.send_log import SendLog
from app.models.task import Task
from app.models.user import User
from app.services import email_sender as es_mod
from app.services.send_engine import get_send_engine
from app.utils.security import encrypt_password
from app.utils.time import utcnow


def _make_sender(email: str) -> int:
    with Session(engine) as session:
        if session.get(User, 1) is None:
            session.add(
                User(
                    id=1,
                    username="concurrency-user",
                    password_hash="test-only",
                    email="concurrency@example.com",
                    role="operator",
                )
            )
            session.commit()
        sender = Sender(
            user_id=1,
            email=email,
            password=encrypt_password("x"),
            smtp_server="smtp.example.com",
            smtp_port=587,
            use_tls=True,
            sender_type="自定义SMTP",
            enabled=True,
            weight=50,
            daily_quota=1000,
            status="active",
        )
        session.add(sender)
        session.commit()
        session.refresh(sender)
        return sender.id


def test_concurrent_send_does_not_crash_or_drop_pending(monkeypatch):
    create_db_and_tables()
    sid = _make_sender(f"conc-{time.time()}@example.com")

    monkeypatch.setattr(
        es_mod.email_sender,
        "send",
        lambda *a, **k: (time.sleep(0.02), (True, ""))[1],
    )

    with Session(engine) as session:
        task = Task(
            user_id=1,
            name="concurrent",
            status="pending",
            sender_ids=json.dumps([sid]),
            recipient_count=12,
            subject="t",
            body="<p>t</p>",
            schedule_type="immediate",
            smart_config=json.dumps(
                {
                    "max_retries": 0,
                    "concurrency_per_sender": 6,
                    "risk_auto_pause_task": False,
                    "batch_size": 50,
                }
            ),
            delay_min=0,
            delay_max=0,
        )
        session.add(task)
        session.commit()
        session.refresh(task)
        tid = task.id
        for i in range(12):
            session.add(
                SendLog(
                    task_id=tid,
                    sender_id=sid,
                    recipient_email=f"u{i}@example.com",
                    recipient_name="n",
                    subject="t",
                    status="pending",
                )
            )
        session.commit()

    get_send_engine().submit(tid)

    final = None
    for _ in range(100):
        time.sleep(0.05)
        with Session(engine) as session:
            task = session.get(Task, tid)
            if task.status in ("completed", "failed"):
                logs = session.exec(select(SendLog).where(SendLog.task_id == tid)).all()
                final = (task.status, task.success_count, task.fail_count, [log.status for log in logs])
                break

    assert final is not None
    status, success, fail, log_status = final
    assert status == "completed"
    assert success == 12
    assert fail == 0
    assert log_status.count("pending") == 0
    assert log_status.count("success") == 12


def test_scheduled_task_auto_starts_without_celery(monkeypatch):
    create_db_and_tables()
    sid = _make_sender(f"sched-{time.time()}@example.com")
    monkeypatch.setattr(es_mod.email_sender, "send", lambda *a, **k: (True, ""))

    with Session(engine) as session:
        task = Task(
            user_id=1,
            name="scheduled-due",
            status="pending",
            sender_ids=json.dumps([sid]),
            recipient_count=1,
            subject="s",
            body="<p>s</p>",
            schedule_type="scheduled",
            schedule_time=utcnow() - timedelta(seconds=1),
            smart_config=json.dumps({"max_retries": 0, "risk_auto_pause_task": False}),
            delay_min=0,
            delay_max=0,
        )
        session.add(task)
        session.commit()
        session.refresh(task)
        tid = task.id
        session.add(
            SendLog(
                task_id=tid,
                sender_id=sid,
                recipient_email="due@example.com",
                recipient_name="n",
                subject="s",
                status="pending",
            )
        )
        session.commit()

    # Force one recovery cycle immediately
    se = get_send_engine()
    with Session(engine) as session:
        se._check_scheduled_due(session)

    final = None
    for _ in range(80):
        time.sleep(0.05)
        with Session(engine) as session:
            task = session.get(Task, tid)
            if task.status in ("completed", "failed", "running", "paused"):
                final = task.status
                if task.status == "completed":
                    break
    assert final in ("completed", "running", "paused")
    with Session(engine) as session:
        task = session.get(Task, tid)
        # eventually should complete
        for _ in range(80):
            if task.status == "completed":
                break
            time.sleep(0.05)
            session.refresh(task)
        assert task.status == "completed"
        assert task.success_count == 1
