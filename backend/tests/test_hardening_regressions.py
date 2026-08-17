import json
import threading
from types import SimpleNamespace
from unittest.mock import patch

import pytest
from pydantic import ValidationError
from sqlmodel import SQLModel, Session, create_engine

from app.api.senders import SenderCreate, create_sender
from app.api.templates import TemplateCreate, create_template
from app.api.tasks import TaskCreate
from app.api.tracking import track_open
from app.api.users import UserCreate
from app.config import settings
from app.models.send_log import SendLog
from app.models.sender import Sender
from app.models.task import Task
from app.models.user import User
from app.services.email_sender import EmailSender, SMTPConnectionPool, _parse_proxy
from app.services.send_engine import SendEngine
from app.utils import security
from app.utils.security import create_tracking_signature, encrypt_password


def test_suite_uses_isolated_application_database():
    assert settings.APP_ENV == "testing"
    assert "email-bulk-tests-" in settings.DATABASE_URL


def test_application_sqlite_enforces_foreign_keys():
    from app.database import engine

    with engine.connect() as connection:
        assert connection.exec_driver_sql("PRAGMA foreign_keys").scalar_one() == 1


def _database(tmp_path, name="hardening.db"):
    database = create_engine(
        f"sqlite:///{tmp_path / name}",
        connect_args={"check_same_thread": False},
    )
    SQLModel.metadata.create_all(database)
    return database


def _runner():
    runner = object.__new__(SendEngine)
    runner._queue = []
    runner._lock = threading.RLock()
    runner._active_tasks = set()
    runner._resume_at = {}
    runner._running = False
    runner._workers = []
    runner._engine_id = "hardening-test"
    runner._stop_event = threading.Event()
    runner._send_slots = threading.BoundedSemaphore(4)
    return runner


def test_explicit_none_security_never_starts_tls():
    pool = SMTPConnectionPool()

    assert pool._uses_starttls(
        SimpleNamespace(smtp_security="none", use_tls=True, smtp_port=587)
    ) is False
    assert pool._uses_starttls(
        SimpleNamespace(smtp_security="starttls", use_tls=False, smtp_port=587)
    ) is True


def test_aliyun_api_never_silently_drops_attachments(monkeypatch):
    from app.services.aliyun_dm import aliyun_dm_sender

    monkeypatch.setattr(
        aliyun_dm_sender,
        "send",
        lambda *args, **kwargs: pytest.fail("Aliyun API must not be called with attachments"),
    )
    result = _runner()._send_one(
        log_id=1,
        sender_snapshot=SimpleNamespace(id=9, sender_type="阿里云邮箱推送"),
        subject="subject",
        body_html="<p>body</p>",
        recipient_email="recipient@example.com",
        recipient_name="Recipient",
        attachments=["attachment.pdf"],
        max_retries=0,
        retry_backoff_base=2.0,
    )

    assert result[1] is False
    assert "attachments" in result[2]


def test_proxy_parser_rejects_unsupported_schemes():
    assert _parse_proxy("ftp://proxy.example.com:21") is None
    assert _parse_proxy("http://proxy.example.com:8080\r\nInjected: yes") is None


def test_attachment_payload_is_cached_between_recipients(tmp_path):
    attachment = tmp_path / "cached.txt"
    attachment.write_bytes(b"cached payload")
    sender = EmailSender()

    with patch("builtins.open", wraps=open) as mocked_open:
        assert sender._read_attachment(str(attachment)) == b"cached payload"
        assert sender._read_attachment(str(attachment)) == b"cached payload"

    assert mocked_open.call_count == 1


def test_old_task_attachment_paths_are_revalidated_before_send(tmp_path, monkeypatch):
    import app.services.send_engine as send_engine_module

    database = _database(tmp_path, "attachments.db")
    monkeypatch.setattr(send_engine_module, "engine", database)
    monkeypatch.setattr(settings, "UPLOAD_DIR", str(tmp_path / "uploads"))

    valid = tmp_path / "uploads" / "1" / "attachments" / "allowed.txt"
    valid.parent.mkdir(parents=True)
    valid.write_text("allowed", encoding="utf-8")
    outside = tmp_path / "server-secret.txt"
    outside.write_text("secret", encoding="utf-8")

    with Session(database) as session:
        sender = Sender(
            user_id=1,
            email="sender@example.com",
            password=encrypt_password("secret"),
            smtp_server="smtp.example.com",
            smtp_port=587,
            enabled=True,
            status="active",
            daily_quota=10,
        )
        session.add(sender)
        session.commit()
        session.refresh(sender)
        task = Task(
            user_id=1,
            name="legacy attachments",
            sender_ids=json.dumps([sender.id]),
            recipient_count=1,
            subject="subject",
            body="<p>body</p>",
            attachments=json.dumps([str(valid), str(outside)]),
            smart_config=json.dumps({"max_retries": 0}),
            delay_min=0,
            delay_max=0,
        )
        session.add(task)
        session.commit()
        session.refresh(task)
        session.add(
            SendLog(
                task_id=task.id,
                sender_id=sender.id,
                recipient_email="recipient@example.com",
                status="pending",
            )
        )
        session.commit()
        task_id = task.id

    captured = []

    def fake_send(*args, **kwargs):
        captured.append(kwargs["attachments"])
        return True, ""

    monkeypatch.setattr(send_engine_module.email_sender, "send", fake_send)
    _runner()._process_task(task_id)

    assert captured == [[str(valid.resolve())]]


def test_template_variables_are_stored_as_json(tmp_path):
    database = _database(tmp_path, "templates.db")
    with Session(database) as session:
        user = User(
            username="template-user",
            password_hash="hash",
            email="template@example.com",
            role="operator",
        )
        session.add(user)
        session.commit()
        session.refresh(user)
        template = create_template(
            TemplateCreate(
                name="Welcome",
                subject="Hello",
                body="<p>Hello</p>",
                variables=["name", "email", "name"],
            ),
            current_user=user,
            session=session,
        )

        assert json.loads(template.variables) == ["name", "email"]


def test_user_create_rejects_unknown_role_and_invalid_email():
    with pytest.raises(ValidationError):
        UserCreate(
            username="valid-user",
            password="long-enough-password",
            email="not-an-email",
            role="superuser",
        )


def test_scheduled_time_is_normalized_to_naive_utc():
    scheduled = TaskCreate(
        name="timezone",
        sender_ids=[1],
        subject="subject",
        body="body",
        recipients=[{"email": "recipient@example.com"}],
        schedule_type="scheduled",
        schedule_time="2026-07-19T20:00:00+08:00",
    )

    assert scheduled.schedule_time.tzinfo is None
    assert scheduled.schedule_time.hour == 12


def test_preset_catalog_entries_are_saveable_and_zero_quota_is_preserved(tmp_path):
    database = _database(tmp_path, "sender-presets.db")
    with Session(database) as session:
        user = User(
            username="preset-user",
            password_hash="hash",
            email="preset@example.com",
            role="operator",
        )
        session.add(user)
        session.commit()
        session.refresh(user)
        sender = create_sender(
            SenderCreate(
                email="relay@example.com",
                password="app-password",
                sender_type="Mail.ru",
                daily_quota=0,
            ),
            current_user=user,
            session=session,
        )

        assert sender.smtp_server == "smtp.mail.ru"
        assert sender.daily_quota == 0


def test_redis_rate_limiter_counts_across_calls(monkeypatch):
    class FakeRedis:
        def __init__(self):
            self.values = {}

        def incr(self, key):
            self.values[key] = self.values.get(key, 0) + 1
            return self.values[key]

        def expire(self, key, seconds):
            return True

        def get(self, key):
            return self.values.get(key)

    fake = FakeRedis()
    monkeypatch.setattr(security, "_get_redis_client", lambda: fake)

    assert security.check_rate_limit("distributed", max_requests=2, window=60)
    assert security.check_rate_limit("distributed", max_requests=2, window=60)
    assert not security.check_rate_limit("distributed", max_requests=2, window=60)
    assert security.get_rate_limit_remaining("distributed", 2, 60) == 0


def test_tracking_open_is_idempotent(tmp_path):
    database = _database(tmp_path, "tracking.db")
    with Session(database) as session:
        task = Task(
            user_id=1,
            name="tracking",
            sender_ids="[]",
            recipient_count=1,
        )
        session.add(task)
        session.commit()
        session.refresh(task)
        log = SendLog(
            task_id=task.id,
            sender_id=0,
            recipient_email="recipient@example.com",
            status="success",
        )
        session.add(log)
        session.commit()
        session.refresh(log)
        signature = create_tracking_signature(log.id, "open")

        track_open(log.id, signature, session)
        session.expire_all()
        track_open(log.id, signature, session)
        session.expire_all()

        assert session.get(Task, task.id).open_count == 1


def test_application_lifespan_releases_resources_on_error(monkeypatch):
    import asyncio

    import app.main as main_module
    import app.services.send_engine as send_engine_module

    events = []

    class FakeResult:
        def all(self):
            return []

    class FakeSession:
        def __init__(self, database):
            self.database = database

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, traceback):
            return False

        def exec(self, statement):
            return FakeResult()

    class FakeDatabase:
        def dispose(self):
            events.append("database-disposed")

    class FakeSendEngine:
        def shutdown(self):
            events.append("send-engine-stopped")

    async def start_dispatcher():
        events.append("dispatcher-started")

    async def stop_dispatcher():
        events.append("dispatcher-stopped")

    monkeypatch.setattr(main_module, "create_db_and_tables", lambda: None)
    monkeypatch.setattr(main_module, "ensure_default_admin", lambda: None)
    monkeypatch.setattr(main_module, "Session", FakeSession)
    monkeypatch.setattr(main_module, "db_engine", FakeDatabase())
    monkeypatch.setattr(main_module, "start_dispatcher", start_dispatcher)
    monkeypatch.setattr(main_module, "stop_dispatcher", stop_dispatcher)
    monkeypatch.setattr(
        send_engine_module,
        "get_send_engine",
        lambda: FakeSendEngine(),
    )

    async def fail_during_lifespan():
        async with main_module.lifespan(main_module.app):
            raise RuntimeError("server failure")

    with pytest.raises(RuntimeError, match="server failure"):
        asyncio.run(fail_during_lifespan())

    assert events == [
        "dispatcher-started",
        "send-engine-stopped",
        "dispatcher-stopped",
        "database-disposed",
    ]


def test_application_lifespan_releases_resources_on_startup_error(monkeypatch):
    import asyncio

    import app.main as main_module
    import app.services.send_engine as send_engine_module

    events = []

    class FailingSession:
        def __init__(self, database):
            self.database = database

        def __enter__(self):
            raise RuntimeError("recovery scan failed")

        def __exit__(self, exc_type, exc, traceback):
            return False

    class FakeDatabase:
        def dispose(self):
            events.append("database-disposed")

    class FakeSendEngine:
        def shutdown(self):
            events.append("send-engine-stopped")

    async def start_dispatcher():
        events.append("dispatcher-started")

    async def stop_dispatcher():
        events.append("dispatcher-stopped")

    monkeypatch.setattr(main_module, "create_db_and_tables", lambda: None)
    monkeypatch.setattr(main_module, "ensure_default_admin", lambda: None)
    monkeypatch.setattr(main_module, "Session", FailingSession)
    monkeypatch.setattr(main_module, "db_engine", FakeDatabase())
    monkeypatch.setattr(main_module, "start_dispatcher", start_dispatcher)
    monkeypatch.setattr(main_module, "stop_dispatcher", stop_dispatcher)
    monkeypatch.setattr(
        send_engine_module,
        "get_send_engine",
        lambda: FakeSendEngine(),
    )

    async def enter_lifespan():
        async with main_module.lifespan(main_module.app):
            pytest.fail("lifespan body must not run after a startup failure")

    with pytest.raises(RuntimeError, match="recovery scan failed"):
        asyncio.run(enter_lifespan())

    assert events == [
        "dispatcher-started",
        "send-engine-stopped",
        "dispatcher-stopped",
        "database-disposed",
    ]
