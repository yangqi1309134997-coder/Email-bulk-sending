from sqlmodel import SQLModel, Field
from datetime import datetime
from typing import Optional


class Sender(SQLModel, table=True):
    __tablename__ = "senders"

    id: Optional[int] = Field(default=None, primary_key=True)
    user_id: int = Field(foreign_key="user.id", index=True)
    email: str = Field(index=True)
    password: str  # AES encrypted
    smtp_server: str
    smtp_port: int = 587
    use_tls: bool = True
    sender_type: str = "自定义SMTP"  # QQ邮箱/163邮箱/Gmail/自定义SMTP
    enabled: bool = True
    weight: int = 50  # 1-100
    daily_quota: int = 500
    daily_sent: int = 0
    success_rate: float = 1.0
    status: str = "active"  # active/paused/banned
    consecutive_failures: int = 0
    paused_until: Optional[datetime] = None
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)

    def is_available(self) -> bool:
        if not self.enabled:
            return False
        if self.status == "banned":
            return False
        if self.daily_sent >= self.daily_quota:
            return False
        if self.status == "paused" and self.paused_until:
            if datetime.utcnow() < self.paused_until:
                return False
            # Pause expired, should be reset
        return True