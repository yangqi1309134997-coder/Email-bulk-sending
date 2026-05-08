from datetime import datetime
from typing import Optional
from fastapi import APIRouter, Depends, HTTPException, status
from pydantic import BaseModel
from sqlmodel import Session, select
from .deps import get_current_user
from ..database import get_session
from ..models.sender import Sender
from ..models.user import User
from ..utils.security import encrypt_password, decrypt_password
import smtplib

router = APIRouter(prefix="/api/senders", tags=["发件人"])


class SenderCreate(BaseModel):
    email: str
    password: str
    smtp_server: str
    smtp_port: int = 587
    use_tls: bool = True
    sender_type: str = "自定义SMTP"
    enabled: bool = True
    weight: int = 50
    daily_quota: int = 500


class SenderUpdate(BaseModel):
    email: Optional[str] = None
    password: Optional[str] = None
    smtp_server: Optional[str] = None
    smtp_port: Optional[int] = None
    use_tls: Optional[bool] = None
    sender_type: Optional[str] = None
    enabled: Optional[bool] = None
    weight: Optional[int] = None
    daily_quota: Optional[int] = None


class SenderResponse(BaseModel):
    id: int
    email: str
    smtp_server: str
    smtp_port: int
    use_tls: bool
    sender_type: str
    enabled: bool
    weight: int
    daily_quota: int
    daily_sent: int
    success_rate: float
    status: str
    created_at: datetime
    updated_at: datetime

    class Config:
        from_attributes = True


@router.get("", response_model=list[SenderResponse])
def list_senders(current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    senders = session.exec(select(Sender).where(Sender.user_id == current_user.id)).all()
    return senders


@router.post("", response_model=SenderResponse)
def create_sender(req: SenderCreate, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    sender = Sender(
        user_id=current_user.id,
        email=req.email,
        password=encrypt_password(req.password),
        smtp_server=req.smtp_server,
        smtp_port=req.smtp_port,
        use_tls=req.use_tls,
        sender_type=req.sender_type,
        enabled=req.enabled,
        weight=req.weight,
        daily_quota=req.daily_quota,
    )
    session.add(sender)
    session.commit()
    session.refresh(sender)
    return sender


@router.put("/{sender_id}", response_model=SenderResponse)
def update_sender(sender_id: int, req: SenderUpdate, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    sender = session.get(Sender, sender_id)
    if not sender or sender.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Sender not found")

    update_data = req.model_dump(exclude_unset=True)
    if "password" in update_data:
        update_data["password"] = encrypt_password(update_data.pop("password"))

    for key, value in update_data.items():
        setattr(sender, key, value)

    sender.updated_at = datetime.utcnow()
    session.add(sender)
    session.commit()
    session.refresh(sender)
    return sender


@router.delete("/{sender_id}")
def delete_sender(sender_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    sender = session.get(Sender, sender_id)
    if not sender or sender.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Sender not found")
    session.delete(sender)
    session.commit()
    return {"message": "Sender deleted"}


@router.post("/{sender_id}/test")
def test_sender(sender_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    sender = session.get(Sender, sender_id)
    if not sender or sender.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Sender not found")

    try:
        pwd = decrypt_password(sender.password)
        server = smtplib.SMTP(sender.smtp_server, sender.smtp_port, timeout=10)
        if sender.use_tls:
            server.starttls()
        server.login(sender.email, pwd)
        server.quit()
        return {"success": True, "message": "SMTP connection successful"}
    except Exception as e:
        return {"success": False, "message": f"Connection failed: {str(e)}"}


@router.post("/{sender_id}/toggle", response_model=SenderResponse)
def toggle_sender(sender_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    sender = session.get(Sender, sender_id)
    if not sender or sender.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Sender not found")
    sender.enabled = not sender.enabled
    sender.updated_at = datetime.utcnow()
    session.add(sender)
    session.commit()
    session.refresh(sender)
    return sender