import json
import threading
import time
from datetime import datetime, timedelta, timezone

import pytest
from sqlmodel import SQLModel, Session, create_engine, select

import app.services.send_engine as send_engine_module
from app.models.send_log import SendLog
from app.models.sender import Sender
from app.models.task import Task
from app.services.send_engine import SendEngine
from app.utils.security import encrypt_password
from app.utils.time import utcnow
from app.utils.time import from_unix_utc, to_unix_utc


@pytest.fixture(autouse=True)
def restore_send_engine_database():
    original_engine = send_engine_module.engine
    yield
    send_engine_module.engine = original_engine


def _database(tmp_path, name):
    db = create_engine(
        f"sqlite:///{tmp_path / name}",
        connect_args={"check_same_thread": False},
    )
    SQLModel.metadata.create_all(db)
    send_engine_module.engine = db
    return db


def _runner():
    runner = object.__new__(SendEngine)
    runner._queue = []
    runner._lock = threading.RLock()
    runner._active_tasks = set()
    runner._resume_at = {}
    runner._running = False
    runner._workers = []
    runner._engine_id = "test-engine"
    return runner


def _seed_task(db, *, recipients=3, quota=100, smart_config=None, sender_status="active"):
    with Session(db) as session:
        sender = Sender(
            user_id=1,
            email="sender@example.com",
            password=encrypt_password("secret"),
            smtp_server="smtp.example.com",
            smtp_port=587,
            use_tls=True,
            sender_type="自定义SMTP",
            enabled=True,
            daily_quota=quota,
            daily_sent=0,
            status=sender_status,
        )
        session.add(sender)
        session.commit()
        session.refresh(sender)

        task = Task(
            user_id=1,
            name="reliability",
            status="pending",
            sender_ids=json.dumps([sender.id]),
            recipient_count=recipients,
            subject="subject",
            body="<p>body</p>",
            smart_config=json.dumps(smart_config or {}),
            delay_min=0,
            delay_max=0,
        )
        session.add(task)
        session.commit()
        session.refresh(task)
        for index in range(recipients):
            session.add(
                SendLog(
                    task_id=task.id,
                    sender_id=sender.id,
                    recipient_email=f"user{index}@example.com",
                    status="pending",
                )
            )
        session.commit()
        return task.id, sender.id


def test_delayed_submit_survives_active_task():
    runner = _runner()
    runner._active_tasks.add(42)

    runner.submit(42, delay_seconds=0.02)
    runner._mark_done(42)
    time.sleep(0.03)

    assert runner._pop_task() == 42


def test_task_lease_is_exclusive_across_engines(tmp_path):
    db = _database(tmp_path, "lease.db")
    task_id, _ = _seed_task(db, recipients=1)
    first = _runner()
    second = _runner()
    second._engine_id = "second-engine"

    assert first._try_acquire_task_lease(task_id) is True
    assert second._try_acquire_task_lease(task_id) is False

    first._release_task_lease(task_id)
    assert second._try_acquire_task_lease(task_id) is True


def test_daily_quota_is_reserved_before_sending(tmp_path, monkeypatch):
    db = _database(tmp_path, "quota.db")
    task_id, sender_id = _seed_task(
        db,
        recipients=5,
        quota=1,
        smart_config={"max_retries": 0, "risk_auto_pause_task": True, "risk_pause_seconds": 60},
    )
    calls = []
    monkeypatch.setattr(
        send_engine_module.email_sender,
        "send",
        lambda *args, **kwargs: (calls.append(kwargs["recipient_email"]), (True, ""))[1],
    )

    _runner()._process_task(task_id)

    with Session(db) as session:
        sender = session.get(Sender, sender_id)
        logs = session.exec(select(SendLog).where(SendLog.task_id == task_id)).all()
        assert len(calls) == 1
        assert sender.daily_sent == 1
        assert sum(log.status == "success" for log in logs) == 1
        assert sum(log.status == "pending" for log in logs) == 4


def test_rate_limited_recipient_remains_pending_for_resume(tmp_path, monkeypatch):
    db = _database(tmp_path, "risk.db")
    task_id, _ = _seed_task(
        db,
        recipients=2,
        smart_config={
            "max_retries": 0,
            "max_consecutive_rate_limits": 1,
            "rate_limit_cooldown": 30,
            "risk_pause_seconds": 30,
            "risk_auto_pause_task": True,
            "auto_resume_after_cooldown": True,
        },
    )
    monkeypatch.setattr(
        send_engine_module.email_sender,
        "send",
        lambda *args, **kwargs: (False, "421 Too many attempts; rate limit exceeded"),
    )

    _runner()._process_task(task_id)

    with Session(db) as session:
        task = session.get(Task, task_id)
        logs = session.exec(select(SendLog).where(SendLog.task_id == task_id)).all()
        assert task.status == "paused"
        assert task.pause_reason == "rate_limit"
        assert task.next_run_at is not None and task.next_run_at > utcnow()
        assert [log.status for log in logs] == ["pending", "pending"]
        assert logs[0].next_attempt_at is not None


def test_refresh_does_not_clear_indefinite_manual_pause():
    paused = Sender(
        id=9,
        user_id=1,
        email="paused@example.com",
        password="x",
        smtp_server="smtp.example.com",
        enabled=True,
        status="paused",
        paused_until=None,
    )

    class FakeSession:
        def get(self, model, sender_id):
            return paused

        def add(self, value):
            pass

        def commit(self):
            pass

        def refresh(self, value):
            pass

    available = _runner()._refresh_senders(FakeSession(), [paused.id])

    assert paused.status == "paused"
    assert available == []


def test_two_senders_can_reach_combined_concurrency_limit(tmp_path, monkeypatch):
    db = _database(tmp_path, "combined-concurrency.db")
    task_id, first_sender_id = _seed_task(
        db,
        recipients=4,
        smart_config={
            "batch_size": 4,
            "concurrency_per_sender": 2,
            "max_retries": 0,
        },
    )
    with Session(db) as session:
        second = Sender(
            user_id=1,
            email="second@example.com",
            password=encrypt_password("secret-2"),
            smtp_server="smtp.example.com",
            smtp_port=587,
            use_tls=True,
            sender_type="自定义SMTP",
            enabled=True,
            daily_quota=100,
            daily_sent=0,
            status="active",
        )
        session.add(second)
        session.commit()
        session.refresh(second)
        second_sender_id = second.id
        task = session.get(Task, task_id)
        task.sender_ids = json.dumps([first_sender_id, second_sender_id])
        session.add(task)
        session.commit()

    lock = threading.Lock()
    active_total = 0
    max_total = 0
    active_by_sender = {first_sender_id: 0, second_sender_id: 0}
    max_by_sender = {first_sender_id: 0, second_sender_id: 0}

    def fake_send(*args, **kwargs):
        nonlocal active_total, max_total
        sender_id = kwargs["sender"].id
        with lock:
            active_total += 1
            active_by_sender[sender_id] += 1
            max_total = max(max_total, active_total)
            max_by_sender[sender_id] = max(
                max_by_sender[sender_id], active_by_sender[sender_id]
            )
        time.sleep(0.08)
        with lock:
            active_total -= 1
            active_by_sender[sender_id] -= 1
        return True, ""

    monkeypatch.setattr(send_engine_module.settings, "MAX_TASK_CONCURRENCY", 4)
    monkeypatch.setattr(send_engine_module.email_sender, "send", fake_send)

    _runner()._process_task(task_id)

    assert max_total == 4
    assert max_by_sender == {first_sender_id: 2, second_sender_id: 2}


def test_recovery_clears_expired_in_memory_deadline_and_requeues(tmp_path):
    db = _database(tmp_path, "expired-recovery.db")
    task_id, _ = _seed_task(db, recipients=1)
    with Session(db) as session:
        task = session.get(Task, task_id)
        task.status = "paused"
        task.pause_reason = "rate_limit"
        task.next_run_at = utcnow() - timedelta(seconds=1)
        session.add(task)
        session.commit()

    runner = _runner()
    runner._resume_at[task_id] = time.time() - 1

    runner._recover_once()

    assert task_id in runner._queue
    assert task_id not in runner._resume_at


def test_process_task_renews_lease_during_long_batch(monkeypatch):
    runner = _runner()
    heartbeats = []
    released = []
    monkeypatch.setattr(runner, "_try_acquire_task_lease", lambda task_id: True)
    monkeypatch.setattr(runner, "_lease_heartbeat_interval", lambda: 0.01)
    monkeypatch.setattr(
        runner,
        "_heartbeat_task_lease",
        lambda task_id: (heartbeats.append(task_id), True)[1],
    )
    monkeypatch.setattr(runner, "_process_task_leased", lambda task_id: time.sleep(0.045))
    monkeypatch.setattr(runner, "_release_task_lease", lambda task_id: released.append(task_id))

    runner._process_task(42)

    assert len(heartbeats) >= 2
    assert released == [42]


def test_sender_runtime_limit_is_shared_across_tasks(monkeypatch):
    runner = _runner()
    runner._sender_slot_limit = 2
    slot, _ = runner._sender_runtime(9)
    assert runner._sender_runtime(9)[0] is slot

    lock = threading.Lock()
    active = 0
    max_active = 0

    def use_sender_slot():
        nonlocal active, max_active
        slot.acquire()
        try:
            with lock:
                active += 1
                max_active = max(max_active, active)
            time.sleep(0.03)
            with lock:
                active -= 1
        finally:
            slot.release()

    threads = [threading.Thread(target=use_sender_slot) for _ in range(6)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()

    assert max_active == 2


def test_smart_schedule_scans_beyond_ineligible_first_page(tmp_path, monkeypatch):
    db = _database(tmp_path, "smart-scan.db")
    task_id, _ = _seed_task(
        db,
        recipients=4,
        smart_config={
            "batch_size": 1,
            "max_retries": 0,
            "auto_resume_after_cooldown": False,
        },
    )
    with Session(db) as session:
        task = session.get(Task, task_id)
        task.schedule_type = "smart"
        session.add(task)
        session.commit()

    delivered = []
    monkeypatch.setattr(
        send_engine_module,
        "_is_good_send_hour",
        lambda email, now=None: email == "user3@example.com",
    )
    monkeypatch.setattr(
        send_engine_module.email_sender,
        "send",
        lambda *args, **kwargs: (delivered.append(kwargs["recipient_email"]), (True, ""))[1],
    )

    _runner()._process_task(task_id)

    assert delivered == ["user3@example.com"]
    with Session(db) as session:
        task = session.get(Task, task_id)
        assert task.status == "paused"
        assert task.pause_reason == "smart_window_wait"


def test_naive_utc_circuit_timestamp_round_trip_ignores_local_timezone():
    stored = datetime(2026, 7, 19, 12, 34, 56)
    timestamp = to_unix_utc(stored)

    assert timestamp == stored.replace(tzinfo=timezone.utc).timestamp()
    assert from_unix_utc(timestamp) == stored
