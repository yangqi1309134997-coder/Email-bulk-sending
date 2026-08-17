from sqlmodel import SQLModel, Field
from datetime import datetime
from typing import Optional
from ..utils.time import utcnow


class Template(SQLModel, table=True):
    __tablename__ = "templates"

    id: Optional[int] = Field(default=None, primary_key=True)
    user_id: int = Field(foreign_key="user.id", index=True)
    name: str
    subject: str = ""
    body: str = ""
    variables: str = "[]"  # JSON array of variable names
    created_at: datetime = Field(default_factory=utcnow)
    updated_at: datetime = Field(default_factory=utcnow)
