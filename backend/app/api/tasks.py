import json
import csv
import io
from datetime import datetime, timezone
from typing import Literal, Optional
from fastapi import APIRouter, Depends, HTTPException, Query, Request, status
from fastapi.responses import StreamingResponse
from pydantic import AliasChoices, BaseModel, ConfigDict, EmailStr, Field, field_validator
from sqlalchemy import func, insert, update
from sqlmodel import Session, select
from .deps import get_current_user
from ..database import get_session
from ..models.task import Task
from ..models.send_log import SendLog
from ..models.sender import Sender
from ..models.user import User
from ..config import settings
from ..utils.security import check_configured_rate_limit
from ..utils.time import utcnow
from .upload import resolve_user_attachment_paths

router = APIRouter(prefix="/api/tasks", tags=["发送任务"])


class RecipientInput(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    email: EmailStr
    name: str = Field(default="", max_length=200)

    @field_validator("name")
    @classmethod
    def validate_name(cls, value: str) -> str:
        if any(ord(char) < 32 or ord(char) == 127 for char in value):
            raise ValueError("Recipient name contains control characters")
        return value


class SmartConfigInput(BaseModel):
    model_config = ConfigDict(extra="forbid", populate_by_name=True)

    max_retries: int = Field(
        default=3, ge=0, le=10,
        validation_alias=AliasChoices("max_retries", "maxRetries"),
    )
    retry_backoff_base: float = Field(
        default=2.0, ge=1.0, le=60.0,
        validation_alias=AliasChoices("retry_backoff_base", "retryBackoffBase"),
    )
    rate_limit_cooldown: int = Field(
        default=60, ge=10, le=3600,
        validation_alias=AliasChoices("rate_limit_cooldown", "rateLimitCooldown"),
    )
    max_consecutive_rate_limits: int = Field(
        default=5, ge=1, le=20,
        validation_alias=AliasChoices(
            "max_consecutive_rate_limits", "maxConsecutiveRateLimits"
        ),
    )
    auto_resume_after_cooldown: bool = Field(
        default=True,
        validation_alias=AliasChoices(
            "auto_resume_after_cooldown", "autoResumeAfterCooldown"
        ),
    )
    risk_auto_pause_task: bool = Field(
        default=True,
        validation_alias=AliasChoices("risk_auto_pause_task", "riskAutoPauseTask"),
    )
    risk_pause_seconds: int = Field(
        default=300, ge=30, le=7200,
        validation_alias=AliasChoices("risk_pause_seconds", "riskPauseSeconds"),
    )
    concurrency_per_sender: int = Field(
        default=3, ge=1, le=20,
        validation_alias=AliasChoices("concurrency_per_sender", "concurrencyPerSender"),
    )
    batch_size: int = Field(
        default=100, ge=1, le=500,
        validation_alias=AliasChoices("batch_size", "batchSize"),
    )
    rate_limit_patterns: list[str] = Field(
        default_factory=lambda: [
            "Too many attempts", "rate limit", "spam", "blocked",
            "too many", "quota", "limit exceeded",
        ],
        min_length=1,
        max_length=50,
        validation_alias=AliasChoices("rate_limit_patterns", "rateLimitPatterns"),
    )

    @field_validator("rate_limit_patterns")
    @classmethod
    def validate_rate_limit_patterns(cls, value: list[str]) -> list[str]:
        cleaned = []
        for pattern in value:
            item = pattern.strip()
            if not item or len(item) > 200 or any(ord(char) < 32 for char in item):
                raise ValueError("Invalid rate-limit pattern")
            cleaned.append(item)
        return list(dict.fromkeys(cleaned))


class TaskCreate(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    name: str = Field(min_length=1, max_length=200)
    sender_ids: list[int] = Field(min_length=1, max_length=100)
    subject: str = Field(min_length=1, max_length=998)
    body: str = Field(min_length=1)
    recipients: list[RecipientInput] = Field(min_length=1, max_length=100000)
    attachments: list[str] = Field(default_factory=list, max_length=50)
    schedule_type: Literal["immediate", "scheduled", "smart"] = "immediate"
    schedule_time: Optional[datetime] = None
    smart_config: SmartConfigInput = Field(default_factory=SmartConfigInput)
    delay_min: int = Field(default=5, ge=0, le=3600)
    delay_max: int = Field(default=15, ge=0, le=3600)
    proxies: list[str] = Field(default_factory=list, max_length=100)
    load_balance_strategy: Literal["round_robin", "weighted", "smart"] = "round_robin"

    @field_validator("name", "subject")
    @classmethod
    def validate_header_text(cls, value: str) -> str:
        if any(ord(char) < 32 or ord(char) == 127 for char in value):
            raise ValueError("Task name and subject cannot contain control characters")
        return value

    @field_validator("sender_ids")
    @classmethod
    def validate_sender_ids(cls, value: list[int]) -> list[int]:
        if any(sender_id <= 0 for sender_id in value):
            raise ValueError("sender_ids must contain positive integers")
        return list(dict.fromkeys(value))

    @field_validator("proxies")
    @classmethod
    def validate_proxies(cls, value: list[str]) -> list[str]:
        from ..services.email_sender import _parse_proxy

        cleaned: list[str] = []
        for raw in value:
            proxy = raw.strip()
            if not proxy or len(proxy) > 2048 or any(ord(char) < 32 for char in proxy):
                raise ValueError("Invalid proxy entry")
            if _parse_proxy(proxy) is None:
                raise ValueError("Invalid proxy entry") from None
            cleaned.append(proxy)
        return list(dict.fromkeys(cleaned))

    @field_validator("schedule_time")
    @classmethod
    def normalize_schedule_time(cls, value: Optional[datetime]) -> Optional[datetime]:
        if value is not None and value.tzinfo is not None:
            return value.astimezone(timezone.utc).replace(tzinfo=None)
        return value


class TaskResponse(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    name: str
    status: str
    sender_ids: list[int]
    recipient_count: int
    success_count: int
    fail_count: int
    open_count: int
    click_count: int
    schedule_type: str
    schedule_time: Optional[datetime] = None
    created_at: datetime
    completed_at: Optional[datetime] = None
    next_run_at: Optional[datetime] = None
    pause_reason: str = ""


class LogResponse(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    sender_id: int
    recipient_email: str
    recipient_name: str
    subject: str
    status: str
    error_message: str
    sent_at: Optional[datetime] = None
    opened_at: Optional[datetime] = None
    clicked_at: Optional[datetime] = None

def task_to_response(task: Task) -> TaskResponse:
    try:
        sender_ids = json.loads(task.sender_ids or "[]")
        if not isinstance(sender_ids, list):
            sender_ids = []
        sender_ids = [int(sender_id) for sender_id in sender_ids if int(sender_id) > 0]
    except (TypeError, ValueError, json.JSONDecodeError):
        sender_ids = []
    return TaskResponse(
        id=task.id,
        name=task.name,
        status=task.status,
        sender_ids=sender_ids,
        recipient_count=task.recipient_count,
        success_count=task.success_count,
        fail_count=task.fail_count,
        open_count=task.open_count,
        click_count=task.click_count,
        schedule_type=task.schedule_type,
        schedule_time=task.schedule_time,
        created_at=task.created_at,
        completed_at=task.completed_at,
        next_run_at=task.next_run_at,
        pause_reason=task.pause_reason,
    )


def task_to_response_with_correction(task: Task, session: Session) -> TaskResponse:
    """Return task response with corrected counts from actual logs."""
    counts = session.exec(
        select(SendLog.status, func.count(SendLog.id))
        .where(SendLog.task_id == task.id)
        .where(SendLog.status.in_(["success", "failed"]))  # type: ignore[attr-defined]
        .group_by(SendLog.status)
    ).all()
    status_counts = {log_status: int(count) for log_status, count in counts}
    actual_success = status_counts.get("success", 0)
    actual_fail = status_counts.get("failed", 0)

    # Fix task counts if drifted
    if task.success_count != actual_success or task.fail_count != actual_fail:
        task.success_count = actual_success
        task.fail_count = actual_fail
        session.add(task)
        session.commit()

    return task_to_response(task)


def _correct_task_counts(tasks: list[Task], session: Session) -> None:
    task_ids = [task.id for task in tasks if task.id is not None]
    if not task_ids:
        return
    rows = session.exec(
        select(SendLog.task_id, SendLog.status, func.count(SendLog.id))
        .where(SendLog.task_id.in_(task_ids))  # type: ignore[attr-defined]
        .where(SendLog.status.in_(["success", "failed"]))  # type: ignore[attr-defined]
        .group_by(SendLog.task_id, SendLog.status)
    ).all()
    counts: dict[int, dict[str, int]] = {}
    for task_id, log_status, count in rows:
        counts.setdefault(task_id, {})[log_status] = int(count)
    dirty = False
    for task in tasks:
        actual = counts.get(task.id, {})
        success = actual.get("success", 0)
        failed = actual.get("failed", 0)
        if task.success_count != success or task.fail_count != failed:
            task.success_count = success
            task.fail_count = failed
            session.add(task)
            dirty = True
    if dirty:
        session.commit()


@router.post("", response_model=TaskResponse)
def create_task(
    req: TaskCreate,
    request: Request,
    current_user: User = Depends(get_current_user),
    session: Session = Depends(get_session),
):
    client_host = request.client.host if request.client else "unknown"
    limit_key = f"task-create:{current_user.id}:{client_host}"
    if not check_configured_rate_limit(limit_key, settings.TASK_RATE_LIMIT):
        raise HTTPException(
            status_code=status.HTTP_429_TOO_MANY_REQUESTS,
            detail="Too many task creation requests",
            headers={"Retry-After": "60"},
        )
    if not req.sender_ids:
        raise HTTPException(status_code=400, detail="至少选择一个发件人")
    if not req.recipients:
        raise HTTPException(status_code=400, detail="收件人列表不能为空")
    if not req.subject or not req.body:
        raise HTTPException(status_code=400, detail="主题和正文不能为空")
    if req.delay_min < 0 or req.delay_max < req.delay_min:
        raise HTTPException(status_code=400, detail="延迟参数不合法")
    if req.schedule_type == "scheduled" and not req.schedule_time:
        raise HTTPException(status_code=400, detail="定时发送必须设置发送时间")
    if req.schedule_type == "scheduled" and req.schedule_time is not None and req.schedule_time <= utcnow():
        raise HTTPException(status_code=400, detail="定时发送时间必须晚于当前时间")
    if len(req.recipients) > int(settings.MAX_RECIPIENTS_PER_TASK):
        raise HTTPException(status_code=400, detail="收件人数量超过系统上限")
    if len(req.body.encode("utf-8")) > int(settings.MAX_TASK_BODY_BYTES):
        raise HTTPException(status_code=413, detail="邮件正文超过系统上限")

    # Validate sender ownership in one query to avoid N+1 database lookups.
    owned_sender_ids = set(
        session.exec(
            select(Sender.id).where(
                Sender.user_id == current_user.id,
                Sender.id.in_(req.sender_ids),  # type: ignore[attr-defined]
            )
        ).all()
    )
    missing_sender_ids = [sid for sid in req.sender_ids if sid not in owned_sender_ids]
    if missing_sender_ids:
        raise HTTPException(
            status_code=400,
            detail=f"Sender {missing_sender_ids[0]} not found or not owned by you",
        )

    # Deduplicate recipients by email
    seen = set()
    clean_recipients = []
    for r in req.recipients:
        email = str(r.email).strip().lower()
        if not email or email in seen:
            continue
        seen.add(email)
        clean_recipients.append({"email": email, "name": (r.name or "").strip()})
    if not clean_recipients:
        raise HTTPException(status_code=400, detail="有效收件人为空")

    try:
        attachment_paths = resolve_user_attachment_paths(current_user.id, req.attachments)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from None

    normalized_smart = req.smart_config.model_dump()

    task = Task(
        user_id=current_user.id,
        name=req.name,
        status="pending",
        sender_ids=json.dumps(req.sender_ids),
        recipient_count=len(clean_recipients),
        subject=req.subject,
        body=req.body,
        attachments=json.dumps(attachment_paths),
        schedule_type=req.schedule_type,
        schedule_time=req.schedule_time,
        smart_config=json.dumps(normalized_smart),
        delay_min=req.delay_min,
        delay_max=req.delay_max,
        proxies=json.dumps(req.proxies or []),
        load_balance_strategy=req.load_balance_strategy or "round_robin",
    )
    try:
        session.add(task)
        session.flush()
        for start in range(0, len(clean_recipients), 1000):
            chunk = clean_recipients[start : start + 1000]
            session.execute(
                insert(SendLog),
                [
                    {
                        "task_id": task.id,
                        "sender_id": req.sender_ids[0],
                        "recipient_email": recipient["email"],
                        "recipient_name": recipient.get("name", ""),
                        "subject": req.subject,
                        "status": "pending",
                    }
                    for recipient in chunk
                ],
            )
        session.commit()
        session.refresh(task)
    except Exception:
        session.rollback()
        raise

    # Dispatch sending via persistent engine for immediate / smart
    if req.schedule_type in ("immediate", "smart"):
        from ..services.send_engine import get_send_engine
        get_send_engine().submit(task.id)

    return task_to_response(task)


@router.get("", response_model=list[TaskResponse])
def list_tasks(
    skip: int = Query(default=0, ge=0),
    limit: int = Query(default=100, ge=1, le=500),
    current_user: User = Depends(get_current_user),
    session: Session = Depends(get_session),
):
    tasks = session.exec(
        select(Task)
        .where(Task.user_id == current_user.id)
        .order_by(Task.created_at.desc())
        .offset(skip)
        .limit(limit)
    ).all()
    _correct_task_counts(tasks, session)
    return [task_to_response(task) for task in tasks]


@router.get("/{task_id}", response_model=TaskResponse)
def get_task(task_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    task = session.get(Task, task_id)
    if not task or task.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Task not found")
    _correct_task_counts([task], session)
    return task_to_response(task)


@router.post("/{task_id}/pause")
def pause_task(task_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    task = session.get(Task, task_id)
    if not task or task.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Task not found")
    if task.status != "running":
        raise HTTPException(status_code=400, detail="Task is not running")
    task.status = "paused"
    task.pause_reason = "manual"
    task.next_run_at = None
    session.add(task)
    session.commit()
    return {"message": "Task paused"}


@router.post("/{task_id}/resume")
def resume_task(task_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    task = session.get(Task, task_id)
    if not task or task.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Task not found")
    if task.status not in ("paused", "running", "failed"):
        raise HTTPException(status_code=400, detail="Task cannot be resumed")
    task.status = "running"
    task.pause_reason = ""
    task.next_run_at = None
    session.add(task)
    session.commit()
    from ..services.send_engine import get_send_engine
    get_send_engine().submit(task.id)
    return {"message": "Task resumed"}


@router.post("/{task_id}/retry-failed")
def retry_failed(task_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    task = session.get(Task, task_id)
    if not task or task.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Task not found")
    if task.status not in ("completed", "failed", "cancelled"):
        raise HTTPException(status_code=400, detail="只能对已完成、失败或已取消的任务重试")

    counts = session.exec(
        select(SendLog.status, func.count(SendLog.id))
        .where(SendLog.task_id == task_id)
        .where(SendLog.status.in_(["success", "failed", "pending"]))  # type: ignore[attr-defined]
        .group_by(SendLog.status)
    ).all()
    status_counts = {log_status: int(count) for log_status, count in counts}
    retryable = status_counts.get("failed", 0) + status_counts.get("pending", 0)
    if not retryable:
        raise HTTPException(status_code=400, detail="没有可重试的记录")

    session.exec(
        update(SendLog)
        .where(SendLog.task_id == task_id, SendLog.status.in_(["failed", "pending"]))  # type: ignore[attr-defined]
        .values(
            status="pending",
            error_message="",
            sent_at=None,
            attempt_count=0,
            claimed_at=None,
            next_attempt_at=None,
            last_error_code="",
        )
        .execution_options(synchronize_session=False)
    )

    task.status = "running"
    task.success_count = status_counts.get("success", 0)
    task.fail_count = 0
    task.completed_at = None
    task.pause_reason = ""
    task.next_run_at = None
    session.add(task)
    session.commit()

    from ..services.send_engine import get_send_engine
    get_send_engine().submit(task.id)
    return {"message": f"已重置 {retryable} 条记录，任务重新开始发送"}


@router.post("/{task_id}/cancel")
def cancel_task(task_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    task = session.get(Task, task_id)
    if not task or task.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Task not found")
    if task.status not in ("pending", "running", "paused"):
        raise HTTPException(status_code=400, detail="Task cannot be cancelled")
    task.status = "cancelled"
    session.add(task)
    session.commit()
    return {"message": "Task cancelled"}


@router.get("/{task_id}/logs", response_model=list[LogResponse])
def get_task_logs(
    task_id: int,
    skip: int = Query(default=0, ge=0),
    limit: int = Query(default=100, ge=1, le=1000),
    current_user: User = Depends(get_current_user),
    session: Session = Depends(get_session),
):
    task = session.get(Task, task_id)
    if not task or task.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Task not found")
    logs = session.exec(
        select(SendLog)
        .where(SendLog.task_id == task_id)
        .order_by(SendLog.id)
        .offset(skip)
        .limit(limit)
    ).all()
    return logs


@router.get("/{task_id}/export")
def export_task_report(task_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    task = session.get(Task, task_id)
    if not task or task.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Task not found")

    def safe_cell(value) -> str:
        text_value = "" if value is None else str(value)
        if text_value.startswith(("=", "+", "-", "@", "\t", "\r")):
            return "'" + text_value
        return text_value

    def stream_rows():
        output = io.StringIO()
        writer = csv.writer(output)

        def render(row) -> str:
            output.seek(0)
            output.truncate(0)
            writer.writerow(row)
            return output.getvalue()

        yield "\ufeff"
        yield render(["序号", "收件人邮箱", "姓名", "状态", "失败原因", "发送时间", "打开时间", "点击时间"])
        index = 0
        last_id = 0
        while True:
            logs = session.exec(
                select(SendLog)
                .where(SendLog.task_id == task_id, SendLog.id > last_id)
                .order_by(SendLog.id)
                .limit(1000)
            ).all()
            if not logs:
                break
            for log in logs:
                index += 1
                yield render([
                    index,
                    safe_cell(log.recipient_email),
                    safe_cell(log.recipient_name),
                    safe_cell(log.status),
                    safe_cell(log.error_message),
                    safe_cell(log.sent_at),
                    safe_cell(log.opened_at),
                    safe_cell(log.clicked_at),
                ])
            last_id = logs[-1].id

    return StreamingResponse(
        stream_rows(),
        media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename=task_{task_id}_report.csv"},
    )
