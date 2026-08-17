from sqlmodel import SQLModel, Field
from datetime import datetime
from typing import Optional
from ..utils.time import utcnow


class SenderTemplate(SQLModel, table=True):
    __tablename__ = "sender_templates"

    id: Optional[int] = Field(default=None, primary_key=True)
    user_id: int = Field(foreign_key="user.id", index=True)
    name: str
    description: str = ""
    sender_type: str
    smtp_server: str
    smtp_port: int
    use_tls: bool = True
    smtp_username: str = ""
    smtp_security: str = ""
    weight: int = 50
    daily_quota: int = 500
    aliyun_access_key: str = ""
    aliyun_access_secret: str = ""  # AES encrypted
    aliyun_region: str = "cn-hangzhou"
    aliyun_from_name: str = ""
    created_at: datetime = Field(default_factory=utcnow)
    updated_at: datetime = Field(default_factory=utcnow)
