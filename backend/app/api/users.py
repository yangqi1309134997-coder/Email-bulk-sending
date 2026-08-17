from datetime import datetime
from typing import Literal, Optional
from fastapi import APIRouter, Depends, HTTPException, Query
from pydantic import BaseModel, ConfigDict, EmailStr, Field, field_validator
from sqlalchemy import func
from sqlalchemy.exc import IntegrityError
from sqlmodel import Session, select
from .deps import require_admin
from ..database import get_session
from ..models.user import User
from ..models.sender import Sender
from ..models.task import Task
from ..models.template import Template
from ..models.sender_template import SenderTemplate
from ..utils.security import get_password_hash

router = APIRouter(prefix="/api/users", tags=["用户管理"])


class UserResponse(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    username: str
    email: str
    role: str
    created_at: datetime
    last_login: Optional[datetime] = None

class UserUpdate(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    role: Optional[Literal["admin", "operator"]] = None
    email: Optional[EmailStr] = None

    @field_validator("email")
    @classmethod
    def normalize_email(cls, value: Optional[EmailStr]) -> Optional[str]:
        return str(value).lower() if value is not None else None


class UserCreate(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    username: str = Field(min_length=3, max_length=64, pattern=r"^[A-Za-z0-9_.-]+$")
    password: str = Field(min_length=10, max_length=128)
    email: EmailStr
    role: Literal["admin", "operator"] = "operator"

    @field_validator("email")
    @classmethod
    def normalize_email(cls, value: EmailStr) -> str:
        return str(value).lower()


@router.get("", response_model=list[UserResponse])
def list_users(
    skip: int = Query(default=0, ge=0),
    limit: int = Query(default=100, ge=1, le=500),
    admin: User = Depends(require_admin),
    session: Session = Depends(get_session),
):
    users = session.exec(
        select(User).order_by(User.id).offset(skip).limit(limit)
    ).all()
    return users


@router.post("", response_model=UserResponse)
def create_user(req: UserCreate, admin: User = Depends(require_admin), session: Session = Depends(get_session)):
    email = str(req.email).lower()
    existing = session.exec(
        select(User).where(
            (func.lower(User.username) == req.username.lower())
            | (func.lower(User.email) == email)
        )
    ).first()
    if existing:
        raise HTTPException(status_code=400, detail="Username or email already exists")

    user = User(
        username=req.username,
        password_hash=get_password_hash(req.password),
        email=email,
        role=req.role,
    )
    try:
        session.add(user)
        session.commit()
    except IntegrityError:
        session.rollback()
        raise HTTPException(status_code=400, detail="Username or email already exists") from None
    session.refresh(user)
    return user


@router.put("/{user_id}", response_model=UserResponse)
def update_user(user_id: int, req: UserUpdate, admin: User = Depends(require_admin), session: Session = Depends(get_session)):
    user = session.get(User, user_id)
    if not user:
        raise HTTPException(status_code=404, detail="User not found")

    update_data = req.model_dump(exclude_unset=True)
    if user.id == admin.id and update_data.get("role") == "operator":
        raise HTTPException(status_code=400, detail="Cannot remove your own admin role")
    if "email" in update_data:
        email = str(update_data["email"]).lower()
        duplicate = session.exec(
            select(User).where(
                User.id != user_id,
                func.lower(User.email) == email,
            )
        ).first()
        if duplicate:
            raise HTTPException(status_code=400, detail="Email already exists")
        update_data["email"] = email
    for key, value in update_data.items():
        setattr(user, key, value)

    try:
        session.add(user)
        session.commit()
    except IntegrityError:
        session.rollback()
        raise HTTPException(status_code=400, detail="User update conflicts with an existing record") from None
    session.refresh(user)
    return user


@router.delete("/{user_id}")
def delete_user(user_id: int, admin: User = Depends(require_admin), session: Session = Depends(get_session)):
    user = session.get(User, user_id)
    if not user:
        raise HTTPException(status_code=404, detail="User not found")
    if user.id == admin.id:
        raise HTTPException(status_code=400, detail="Cannot delete yourself")
    if user.role == "admin" and len(
        session.exec(select(User).where(User.role == "admin").limit(2)).all()
    ) <= 1:
        raise HTTPException(status_code=400, detail="Cannot delete the last administrator")
    owned_data = (
        session.exec(select(Sender.id).where(Sender.user_id == user_id).limit(1)).first()
        or session.exec(select(Task.id).where(Task.user_id == user_id).limit(1)).first()
        or session.exec(select(Template.id).where(Template.user_id == user_id).limit(1)).first()
        or session.exec(select(SenderTemplate.id).where(SenderTemplate.user_id == user_id).limit(1)).first()
    )
    if owned_data is not None:
        raise HTTPException(
            status_code=400,
            detail="Cannot delete a user with senders, tasks, or templates",
        )
    session.delete(user)
    session.commit()
    return {"message": "User deleted"}
