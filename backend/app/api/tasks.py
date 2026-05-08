import json
from datetime import datetime
from typing import Optional
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from sqlmodel import Session, select
from .deps import get_current_user
from ..database import get_session
from ..models.task import Task
from ..models.send_log import SendLog
from ..models.sender import Sender
from ..models.user import User
from ..tasks.send_email import send_batch_task

router = APIRouter(prefix="/api/tasks", tags=["发送任务"])


class TaskCreate(BaseModel):
    name: str
    sender_ids: list[int]
    subject: str
    body: str
    recipients: list[dict]  # [{"email": "", "name": ""}]
    attachments: list[str] = []
    schedule_type: str = "immediate"
    schedule_time: Optional[datetime] = None
    smart_config: Optional[dict] = None
    delay_min: int = 5
    delay_max: int = 15
    proxies: list[str] = []
    load_balance_strategy: str = "round_robin"


class TaskResponse(BaseModel):
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

    class Config:
        from_attributes = True


class LogResponse(BaseModel):
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

    class Config:
        from_attributes = True


def task_to_response(task: Task) -> TaskResponse:
    return TaskResponse(
        id=task.id,
        name=task.name,
        status=task.status,
        sender_ids=json.loads(task.sender_ids),
        recipient_count=task.recipient_count,
        success_count=task.success_count,
        fail_count=task.fail_count,
        open_count=task.open_count,
        click_count=task.click_count,
        schedule_type=task.schedule_type,
        schedule_time=task.schedule_time,
        created_at=task.created_at,
        completed_at=task.completed_at,
    )


@router.post("", response_model=TaskResponse)
def create_task(req: TaskCreate, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    # Validate sender ownership
    for sid in req.sender_ids:
        sender = session.get(Sender, sid)
        if not sender or sender.user_id != current_user.id:
            raise HTTPException(status_code=400, detail=f"Sender {sid} not found or not owned by you")

    task = Task(
        user_id=current_user.id,
        name=req.name,
        status="pending",
        sender_ids=json.dumps(req.sender_ids),
        recipient_count=len(req.recipients),
        subject=req.subject,
        body=req.body,
        attachments=json.dumps(req.attachments),
        schedule_type=req.schedule_type,
        schedule_time=req.schedule_time,
        smart_config=json.dumps(req.smart_config or {}),
        delay_min=req.delay_min,
        delay_max=req.delay_max,
        proxies=json.dumps(req.proxies),
        load_balance_strategy=req.load_balance_strategy,
    )
    session.add(task)
    session.commit()
    session.refresh(task)

    # Store recipients as send_logs with pending status
    for r in req.recipients:
        log = SendLog(
            task_id=task.id,
            sender_id=req.sender_ids[0],
            recipient_email=r.get("email", ""),
            recipient_name=r.get("name", ""),
            subject=req.subject,
            status="pending",
        )
        session.add(log)
    session.commit()

    # Dispatch Celery task for immediate sending
    if req.schedule_type == "immediate":
        send_batch_task.delay(task.id)

    return task_to_response(task)


@router.get("", response_model=list[TaskResponse])
def list_tasks(current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    tasks = session.exec(select(Task).where(Task.user_id == current_user.id).order_by(Task.created_at.desc())).all()
    return [task_to_response(t) for t in tasks]


@router.get("/{task_id}", response_model=TaskResponse)
def get_task(task_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    task = session.get(Task, task_id)
    if not task or task.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Task not found")
    return task_to_response(task)


@router.post("/{task_id}/pause")
def pause_task(task_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    task = session.get(Task, task_id)
    if not task or task.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Task not found")
    if task.status != "running":
        raise HTTPException(status_code=400, detail="Task is not running")
    task.status = "paused"
    session.add(task)
    session.commit()
    return {"message": "Task paused"}


@router.post("/{task_id}/resume")
def resume_task(task_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    task = session.get(Task, task_id)
    if not task or task.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Task not found")
    if task.status != "paused":
        raise HTTPException(status_code=400, detail="Task is not paused")
    task.status = "running"
    session.add(task)
    session.commit()
    return {"message": "Task resumed"}


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
def get_task_logs(task_id: int, skip: int = 0, limit: int = 100, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    task = session.get(Task, task_id)
    if not task or task.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Task not found")
    logs = session.exec(
        select(SendLog).where(SendLog.task_id == task_id).offset(skip).limit(limit)
    ).all()
    return logs


@router.get("/{task_id}/export")
def export_task_report(task_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    import csv
    import io
    from fastapi.responses import StreamingResponse

    task = session.get(Task, task_id)
    if not task or task.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Task not found")

    logs = session.exec(select(SendLog).where(SendLog.task_id == task_id)).all()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["序号", "收件人邮箱", "姓名", "状态", "失败原因", "发送时间", "打开时间", "点击时间"])
    for i, log in enumerate(logs, 1):
        writer.writerow([
            i, log.recipient_email, log.recipient_name, log.status,
            log.error_message, log.sent_at, log.opened_at, log.clicked_at,
        ])

    output.seek(0)
    return StreamingResponse(
        io.BytesIO(output.getvalue().encode("utf-8-sig")),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename=task_{task_id}_report.csv"},
    )