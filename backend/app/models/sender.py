from sqlmodel import SQLModel, Field
from datetime import datetime
from typing import Optional
from ..utils.time import utcnow


class Sender(SQLModel, table=True):
    __tablename__ = "senders"

    id: Optional[int] = Field(default=None, primary_key=True)
    user_id: int = Field(foreign_key="user.id", index=True)
    email: str = Field(index=True)
    password: str  # AES encrypted
    smtp_server: str
    smtp_port: int = 587
    use_tls: bool = True
    smtp_username: str = ""
    smtp_security: str = ""  # ssl/starttls/none; empty derives from port/use_tls
    sender_type: str = "自定义SMTP"  # QQ邮箱/163邮箱/Gmail/阿里云邮箱推送/自定义SMTP
    enabled: bool = True
    weight: int = 50  # 1-100
    daily_quota: int = 500
    daily_sent: int = 0
    success_rate: float = 1.0
    status: str = "active"  # active/paused/banned
    consecutive_failures: int = 0
    paused_until: Optional[datetime] = None
    # 阿里云邮箱推送专属字段
    aliyun_access_key: str = ""
    aliyun_access_secret: str = ""
    aliyun_region: str = "cn-hangzhou"
    aliyun_from_name: str = ""
    created_at: datetime = Field(default_factory=utcnow)
    updated_at: datetime = Field(default_factory=utcnow)

    # Circuit breaker fields (persisted)
    cb_state: str = "closed"  # closed, open, half-open
    cb_failure_count: int = 0
    cb_success_count: int = 0
    cb_next_attempt_time: Optional[datetime] = None
    cb_last_failure_time: Optional[datetime] = None

    def is_available(self) -> bool:
        """Whether this sender can be selected for sending right now."""
        if not self.enabled:
            return False
        if self.status == "banned":
            return False
        daily_quota = self.daily_quota or 0
        daily_sent = self.daily_sent or 0
        if daily_quota > 0 and daily_sent >= daily_quota:
            return False
        if self.status == "paused":
            # Paused without expiry stays unavailable until explicitly unpaused
            if self.paused_until is None:
                return False
            if utcnow() < self.paused_until:
                return False
            return True
        return self.status in ("active", "")
