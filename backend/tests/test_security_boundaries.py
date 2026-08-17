import hashlib
import hmac
import re
import base64
from datetime import timedelta
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

import pytest
from fastapi.testclient import TestClient
from sqlmodel import SQLModel, Session, create_engine, select
from starlette.websockets import WebSocketDisconnect

import app.main as main_module
from app.config import settings
from app.config import Settings
from app.database import get_session
from app.models.send_log import SendLog
from app.models.sender import Sender
from app.models.task import Task
from app.models.user import User
from app.services.tracker import replace_links_with_tracking
from app.utils import security as security_module
from app.utils.security import (
    create_access_token,
    decrypt_password,
    encrypt_password,
    get_password_hash,
)
from app.utils.time import utcnow
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad


@pytest.fixture
def api(tmp_path, monkeypatch):
    db = create_engine(
        f"sqlite:///{tmp_path / 'security.db'}",
        connect_args={"check_same_thread": False},
    )
    SQLModel.metadata.create_all(db)

    def override_session():
        with Session(db) as session:
            yield session

    main_module.app.dependency_overrides[get_session] = override_session
    monkeypatch.setattr(main_module, "db_engine", db)
    monkeypatch.setattr(settings, "UPLOAD_DIR", str(tmp_path / "uploads"))
    security_module._rate_limit_store.clear()
    client = TestClient(main_module.app)
    try:
        yield client, db, tmp_path
    finally:
        client.close()
        main_module.app.dependency_overrides.clear()
        security_module._rate_limit_store.clear()
        db.dispose()


def _seed_user(db, *, username="operator", role="operator"):
    with Session(db) as session:
        user = User(
            username=username,
            password_hash=get_password_hash("correct horse battery staple"),
            email=f"{username}@example.com",
            role=role,
        )
        session.add(user)
        session.commit()
        session.refresh(user)
        return user.id


def _auth(user_id):
    return {"Authorization": f"Bearer {create_access_token({'sub': user_id})}"}


def _tracking_signature(log_id: int, action: str) -> str:
    payload = f"{action}:{log_id}".encode("utf-8")
    return hmac.new(settings.SECRET_KEY.encode("utf-8"), payload, hashlib.sha256).hexdigest()


def _seed_task_log(db, user_id):
    with Session(db) as session:
        task = Task(
            user_id=user_id,
            name="tracking",
            status="completed",
            sender_ids="[]",
            recipient_count=1,
            subject="subject",
            body="body",
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
        return task.id, log.id


def test_public_registration_cannot_assign_admin_role(api):
    client, db, _ = api

    response = client.post(
        "/api/auth/register",
        json={
            "username": "attacker",
            "password": "long-enough-password",
            "email": "attacker@example.com",
            "role": "admin",
        },
    )

    assert response.status_code == 422
    with Session(db) as session:
        assert session.exec(select(User).where(User.username == "attacker")).first() is None


def test_login_endpoint_enforces_configured_rate_limit(api, monkeypatch):
    client, _, _ = api
    monkeypatch.setattr(settings, "LOGIN_RATE_LIMIT", "2/minute")

    responses = [
        client.post("/api/auth/login", json={"username": "missing", "password": "invalid"})
        for _ in range(3)
    ]

    assert [response.status_code for response in responses] == [401, 401, 429]


def test_task_rejects_attachment_outside_current_users_upload_root(api):
    client, db, tmp_path = api
    user_id = _seed_user(db)
    secret_path = tmp_path / "server-secret.txt"
    secret_path.write_text("do not attach", encoding="utf-8")
    with Session(db) as session:
        sender = Sender(
            user_id=user_id,
            email="sender@example.com",
            password=encrypt_password("secret"),
            smtp_server="smtp.example.com",
            smtp_port=587,
            enabled=True,
            status="active",
        )
        session.add(sender)
        session.commit()
        session.refresh(sender)
        sender_id = sender.id

    response = client.post(
        "/api/tasks",
        headers=_auth(user_id),
        json={
            "name": "invalid attachment",
            "sender_ids": [sender_id],
            "subject": "subject",
            "body": "<p>body</p>",
            "recipients": [{"email": "recipient@example.com", "name": "Recipient"}],
            "attachments": [str(secret_path)],
            "schedule_type": "scheduled",
            "schedule_time": (utcnow() + timedelta(hours=1)).isoformat(),
        },
    )

    assert response.status_code == 400
    with Session(db) as session:
        assert session.exec(select(Task)).all() == []


def test_task_rejects_malformed_recipient_email(api):
    client, db, _ = api
    user_id = _seed_user(db)
    with Session(db) as session:
        sender = Sender(
            user_id=user_id,
            email="sender@example.com",
            password=encrypt_password("secret"),
            smtp_server="smtp.example.com",
            smtp_port=587,
            enabled=True,
            status="active",
        )
        session.add(sender)
        session.commit()
        session.refresh(sender)
        sender_id = sender.id

    response = client.post(
        "/api/tasks",
        headers=_auth(user_id),
        json={
            "name": "invalid recipient",
            "sender_ids": [sender_id],
            "subject": "subject",
            "body": "<p>body</p>",
            "recipients": [{"email": "not-an-email", "name": "Recipient"}],
            "schedule_type": "scheduled",
            "schedule_time": (utcnow() + timedelta(hours=1)).isoformat(),
        },
    )

    assert response.status_code == 422


def test_task_creation_bulk_inserts_deduplicated_recipients_atomically(api):
    client, db, _ = api
    user_id = _seed_user(db)
    with Session(db) as session:
        sender = Sender(
            user_id=user_id,
            email="sender@example.com",
            password=encrypt_password("secret"),
            smtp_server="smtp.example.com",
            smtp_port=587,
            enabled=True,
            status="active",
        )
        session.add(sender)
        session.commit()
        session.refresh(sender)
        sender_id = sender.id

    response = client.post(
        "/api/tasks",
        headers=_auth(user_id),
        json={
            "name": "bulk insert",
            "sender_ids": [sender_id],
            "subject": "subject",
            "body": "<p>body</p>",
            "recipients": [
                {"email": "one@example.com", "name": "One"},
                {"email": "ONE@example.com", "name": "Duplicate"},
                {"email": "two@example.com", "name": "Two"},
            ],
            "schedule_type": "scheduled",
            "schedule_time": (utcnow() + timedelta(hours=1)).isoformat(),
        },
    )

    assert response.status_code == 200
    task_id = response.json()["id"]
    with Session(db) as session:
        task = session.get(Task, task_id)
        logs = session.exec(select(SendLog).where(SendLog.task_id == task_id)).all()
        assert task.recipient_count == 2
        assert [log.recipient_email for log in logs] == [
            "one@example.com",
            "two@example.com",
        ]


def test_upload_sanitizes_filename_and_scopes_file_to_user(api):
    client, db, _ = api
    user_id = _seed_user(db)

    response = client.post(
        "/api/upload/attachment",
        headers=_auth(user_id),
        files={"file": ("../../evil.txt", b"content", "text/plain")},
    )

    assert response.status_code == 200
    stored = Path(response.json()["path"]).resolve()
    expected_root = (Path(settings.UPLOAD_DIR) / str(user_id) / "attachments").resolve()
    assert stored.is_relative_to(expected_root)
    assert stored.name.endswith("_evil.txt")
    assert stored.read_bytes() == b"content"


def test_oversized_upload_removes_partial_file(api, monkeypatch):
    client, db, _ = api
    user_id = _seed_user(db)
    monkeypatch.setattr(settings, "MAX_ATTACHMENT_SIZE_MB", 1)

    response = client.post(
        "/api/upload/attachment",
        headers=_auth(user_id),
        files={"file": ("large.txt", b"x" * (1024 * 1024 + 1), "text/plain")},
    )

    assert response.status_code == 413
    upload_root = Path(settings.UPLOAD_DIR) / str(user_id) / "attachments"
    assert not upload_root.exists() or list(upload_root.iterdir()) == []


def test_encrypted_secret_detects_ciphertext_tampering():
    encrypted = encrypt_password("smtp-secret")
    assert encrypted.startswith("v3:")
    payload = bytearray(base64.b64decode(encrypted[3:]))
    payload[-1] ^= 1
    tampered = "v3:" + base64.b64encode(payload).decode("ascii")

    with pytest.raises(ValueError, match="Invalid encrypted secret"):
        decrypt_password(tampered)


def test_legacy_v2_secret_remains_readable():
    key = settings.AES_KEY.encode("utf-8")[:32]
    iv = b"legacy-v2-vector"
    cipher = AES.new(key, AES.MODE_CBC, iv=iv)
    ciphertext = cipher.encrypt(pad(b"legacy-secret", AES.block_size))
    encrypted = "v2:" + base64.b64encode(iv + ciphertext).decode("ascii")

    assert decrypt_password(encrypted) == "legacy-secret"


def test_production_settings_reject_default_secrets():
    with pytest.raises(ValueError, match="production secrets"):
        Settings(
            APP_ENV="production",
            SECRET_KEY="change-me-to-a-random-secret-key-at-least-32-chars",
            AES_KEY="change-me-to-32-bytes-aes-key-123456",
        )


def test_unsigned_tracking_request_is_rejected(api):
    client, db, _ = api
    user_id = _seed_user(db)
    task_id, log_id = _seed_task_log(db, user_id)

    response = client.get(f"/track/open/{log_id}")

    assert response.status_code == 403
    with Session(db) as session:
        assert session.get(SendLog, log_id).opened_at is None
        assert session.get(Task, task_id).open_count == 0


def test_tracking_link_preserves_full_url_and_adds_signature():
    original = "https://example.com/path?a=1&b=two%20words#section"

    rewritten = replace_links_with_tracking(f'<a href="{original}">open</a>', 17)

    href = re.search(r'href=["\']([^"\']+)', rewritten).group(1).replace("&amp;", "&")
    query = parse_qs(urlsplit(href).query)
    assert query["url"] == [original]
    assert query["sig"] == [_tracking_signature(17, "click")]


def test_private_redirect_is_rejected_before_click_is_recorded(api):
    client, db, _ = api
    user_id = _seed_user(db)
    task_id, log_id = _seed_task_log(db, user_id)

    response = client.get(
        f"/track/click/{log_id}",
        params={
            "url": "http://127.0.0.2/internal",
            "sig": _tracking_signature(log_id, "click"),
        },
        follow_redirects=False,
    )

    assert response.status_code == 400
    with Session(db) as session:
        assert session.get(SendLog, log_id).clicked_at is None
        assert session.get(Task, task_id).click_count == 0


def test_websocket_requires_token_and_task_ownership(api):
    client, db, _ = api
    owner_id = _seed_user(db, username="owner")
    stranger_id = _seed_user(db, username="stranger")
    task_id, _ = _seed_task_log(db, owner_id)

    with pytest.raises(WebSocketDisconnect) as missing:
        with client.websocket_connect(f"/ws/tasks/{task_id}"):
            pass
    assert missing.value.code == 4401

    stranger_token = create_access_token({"sub": stranger_id})
    with pytest.raises(WebSocketDisconnect) as forbidden:
        with client.websocket_connect(f"/ws/tasks/{task_id}?token={stranger_token}"):
            pass
    assert forbidden.value.code == 4403

    owner_token = create_access_token({"sub": owner_id})
    with client.websocket_connect(
        f"/ws/tasks/{task_id}",
        subprotocols=[f"access-token.{owner_token}"],
    ) as websocket:
        assert websocket.accepted_subprotocol == f"access-token.{owner_token}"

    # Legacy query-string clients remain accepted during migration.
    with client.websocket_connect(f"/ws/tasks/{task_id}?token={owner_token}"):
        pass
