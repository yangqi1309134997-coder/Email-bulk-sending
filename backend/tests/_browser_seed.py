import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from sqlmodel import Session, SQLModel, create_engine, select

from app.models import *  # noqa: F401,F403
from app.models.sender import Sender
from app.models.user import User
from app.utils.security import encrypt_password, get_password_hash
from app.config import settings

engine = create_engine(settings.DATABASE_URL, connect_args={"check_same_thread": False})
SQLModel.metadata.create_all(engine)
with Session(engine) as session:
    user = session.exec(select(User).where(User.username == "browser-admin")).first()
    if not user:
        user = User(
            username="browser-admin",
            password_hash=get_password_hash("correct horse battery staple"),
            email="browser-admin@example.com",
            role="admin",
        )
        session.add(user)
        session.commit()
        session.refresh(user)
    sender = session.exec(select(Sender).where(Sender.email == "browser-sender@example.com")).first()
    if not sender:
        sender = Sender(
            user_id=user.id,
            email="browser-sender@example.com",
            password=encrypt_password("not-used-in-scheduled-test"),
            smtp_server="smtp.example.invalid",
            smtp_port=587,
            use_tls=True,
            smtp_security="starttls",
            enabled=True,
            status="active",
            daily_quota=100,
        )
        session.add(sender)
        session.commit()
    user_id = user.id
    sender_id = sender.id
print(user_id, sender_id)
