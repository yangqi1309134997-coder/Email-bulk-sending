from sqlmodel import SQLModel, Field
from datetime import datetime
from typing import Optional


class SendLog(SQLModel, table=True):
    __tablename__ = "send_logs"

    id: Optional[int] = Field(default=None, primary_key=True)
    task_id: int = Field(foreign_key="send_tasks.id", index=True)
    sender_id: int = Field(foreign_key="senders.id")
    recipient_email: str = Field(index=True)
    recipient_name: str = ""
    subject: str = ""
    status: str = "pending"  # pending/success/failed/bounced
    error_message: str = ""
    sent_at: Optional[datetime] = None
    opened_at: Optional[datetime] = None
    clicked_at: Optional[datetime] = None