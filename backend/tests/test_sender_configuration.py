import json

import pytest
from fastapi import HTTPException
from sqlmodel import Session, SQLModel, create_engine

from app.api.senders import (
    ApplyTemplateRequest,
    SenderCreate,
    SenderTemplateCreate,
    SenderUpdate,
    apply_sender_template,
    create_sender,
    create_sender_from_preset,
    create_sender_template,
    delete_sender,
    delete_sender_template,
    update_sender,
)
from app.models.send_log import SendLog
from app.models.sender import Sender
from app.models.task import Task
from app.models.user import User
from app.utils.security import decrypt_password


def _db(tmp_path):
    db = create_engine(
        f"sqlite:///{tmp_path / 'sender-config.db'}",
        connect_args={"check_same_thread": False},
    )
    SQLModel.metadata.create_all(db)
    return db


def _user(db):
    with Session(db) as session:
        user = User(
            username="config-user",
            password_hash="hash",
            email="config@example.com",
            role="operator",
        )
        session.add(user)
        session.commit()
        session.refresh(user)
        return user


def test_custom_smtp_username_and_security_are_persisted(tmp_path):
    db = _db(tmp_path)
    user = _user(db)
    with Session(db) as session:
        sender = create_sender(
            SenderCreate(
                email="from@example.com",
                password="app-password",
                smtp_server="smtp.example.com",
                smtp_port=465,
                use_tls=True,
                smtp_username="relay-user",
                smtp_security="ssl",
                sender_type="自定义SMTP",
            ),
            current_user=user,
            session=session,
        )
        assert sender.smtp_username == "relay-user"
        assert sender.smtp_security == "ssl"
        assert decrypt_password(sender.password) == "app-password"


def test_aliyun_dm_preset_creates_api_sender(tmp_path):
    db = _db(tmp_path)
    user = _user(db)
    with Session(db) as session:
        sender = create_sender_from_preset(
            "aliyun_dm",
            SenderCreate(
                email="verified@example.com",
                password="",
                sender_type="阿里云邮箱推送",
                aliyun_access_key="access-key",
                aliyun_access_secret="access-secret",
            ),
            current_user=user,
            session=session,
        )
        assert sender.sender_type == "阿里云邮箱推送"
        assert sender.smtp_server == "dm.aliyuncs.com"
        assert sender.smtp_port == 0
        assert decrypt_password(sender.password) == "access-secret"


def test_sender_template_round_trip_requires_smtp_password(tmp_path):
    db = _db(tmp_path)
    user = _user(db)
    with Session(db) as session:
        template = create_sender_template(
            SenderTemplateCreate(
                name="Gmail relay",
                sender_type="Gmail",
                smtp_server="smtp.gmail.com",
                smtp_port=587,
                smtp_security="starttls",
            ),
            current_user=user,
            session=session,
        )
        with pytest.raises(HTTPException, match="密码"):
            apply_sender_template(
                template.id,
                ApplyTemplateRequest(email="new@example.com"),
                current_user=user,
                session=session,
            )

        sender = apply_sender_template(
            template.id,
            ApplyTemplateRequest(email="new@example.com", password="app-password"),
            current_user=user,
            session=session,
        )
        assert sender.smtp_server == "smtp.gmail.com"
        assert decrypt_password(sender.password) == "app-password"


def test_aliyun_template_reuses_encrypted_secret_and_delete_is_safe(tmp_path):
    db = _db(tmp_path)
    user = _user(db)
    with Session(db) as session:
        template = create_sender_template(
            SenderTemplateCreate(
                name="Aliyun API",
                sender_type="阿里云邮箱推送",
                smtp_server="",
                smtp_port=0,
                aliyun_access_key="access-key",
                aliyun_access_secret="access-secret",
            ),
            current_user=user,
            session=session,
        )
        sender = apply_sender_template(
            template.id,
            ApplyTemplateRequest(email="aliyun@example.com"),
            current_user=user,
            session=session,
        )
        assert sender.sender_type == "阿里云邮箱推送"
        assert decrypt_password(sender.password) == "access-secret"
        assert delete_sender_template(template.id, current_user=user, session=session)["message"]


def test_sender_update_rejects_duplicate_email_and_invalid_transport(tmp_path):
    db = _db(tmp_path)
    user = _user(db)
    with Session(db) as session:
        first = create_sender(
            SenderCreate(
                email="first@example.com",
                password="secret",
                smtp_server="smtp.example.com",
                smtp_port=587,
                sender_type="自定义SMTP",
            ),
            current_user=user,
            session=session,
        )
        create_sender(
            SenderCreate(
                email="second@example.com",
                password="secret",
                smtp_server="smtp.example.com",
                smtp_port=587,
                sender_type="自定义SMTP",
            ),
            current_user=user,
            session=session,
        )
        with pytest.raises(HTTPException, match="已存在"):
            update_sender(
                first.id,
                SenderUpdate(email="second@example.com"),
                current_user=user,
                session=session,
            )
        with pytest.raises(HTTPException, match="服务器和端口"):
            update_sender(
                first.id,
                SenderUpdate(smtp_port=0),
                current_user=user,
                session=session,
            )


def test_sender_with_send_history_cannot_be_deleted(tmp_path):
    db = _db(tmp_path)
    user = _user(db)
    with Session(db) as session:
        sender = create_sender(
            SenderCreate(
                email="history@example.com",
                password="secret",
                smtp_server="smtp.example.com",
                smtp_port=587,
                sender_type="自定义SMTP",
            ),
            current_user=user,
            session=session,
        )
        task = Task(
            user_id=user.id,
            name="completed history",
            status="completed",
            sender_ids=json.dumps([sender.id]),
            recipient_count=1,
        )
        session.add(task)
        session.flush()
        session.add(
            SendLog(
                task_id=task.id,
                sender_id=sender.id,
                recipient_email="recipient@example.com",
                status="success",
            )
        )
        session.commit()

        with pytest.raises(HTTPException) as error:
            delete_sender(sender.id, current_user=user, session=session)

        assert error.value.status_code == 409
        assert session.get(Sender, sender.id) is not None


def test_sender_used_by_pending_task_cannot_be_deleted_without_logs(tmp_path):
    db = _db(tmp_path)
    user = _user(db)
    with Session(db) as session:
        sender = create_sender(
            SenderCreate(
                email="pending@example.com",
                password="secret",
                smtp_server="smtp.example.com",
                smtp_port=587,
                sender_type="自定义SMTP",
            ),
            current_user=user,
            session=session,
        )
        session.add(
            Task(
                user_id=user.id,
                name="pending task",
                status="pending",
                sender_ids=json.dumps([sender.id]),
                recipient_count=0,
            )
        )
        session.commit()

        with pytest.raises(HTTPException) as error:
            delete_sender(sender.id, current_user=user, session=session)

        assert error.value.status_code == 409
        assert session.get(Sender, sender.id) is not None


def test_unused_sender_can_be_deleted(tmp_path):
    db = _db(tmp_path)
    user = _user(db)
    with Session(db) as session:
        sender = create_sender(
            SenderCreate(
                email="unused@example.com",
                password="secret",
                smtp_server="smtp.example.com",
                smtp_port=587,
                sender_type="自定义SMTP",
            ),
            current_user=user,
            session=session,
        )
        sender_id = sender.id

        assert delete_sender(sender_id, current_user=user, session=session)["message"]
        assert session.get(Sender, sender_id) is None
