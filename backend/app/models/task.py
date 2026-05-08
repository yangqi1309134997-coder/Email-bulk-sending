from sqlmodel import SQLModel, Field
from datetime import datetime
from typing import Optional


class Task(SQLModel, table=True):
    __tablename__ = "send_tasks"

    id: Optional[int] = Field(default=None, primary_key=True)
    user_id: int = Field(foreign_key="user.id", index=True)
    name: str
    status: str = "pending"  # pending/running/paused/completed/cancelled
    sender_ids: str = "[]"  # JSON array
    recipient_count: int = 0
    success_count: int = 0
    fail_count: int = 0
    open_count: int = 0
    click_count: int = 0
    subject: str = ""
    body: str = ""
    attachments: str = "[]"  # JSON array of file paths
    schedule_type: str = "immediate"  # immediate/scheduled/smart
    schedule_time: Optional[datetime] = None
    smart_config: str = "{}"  # JSON
    delay_min: int = 5
    delay_max: int = 15
    proxies: str = "[]"  # JSON array
    load_balance_strategy: str = "round_robin"  # round_robin/weighted/smart
    created_at: datetime = Field(default_factory=datetime.utcnow)
    completed_at: Optional[datetime] = None