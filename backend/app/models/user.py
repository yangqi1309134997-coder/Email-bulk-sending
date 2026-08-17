from sqlmodel import SQLModel, Field
from datetime import datetime
from typing import Optional
from ..utils.time import utcnow


class User(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    username: str = Field(unique=True, index=True)
    password_hash: str
    role: str = Field(default="operator")  # admin / operator
    email: str
    created_at: datetime = Field(default_factory=utcnow)
    last_login: Optional[datetime] = None
