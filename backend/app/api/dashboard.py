from datetime import datetime, timedelta
from fastapi import APIRouter, Depends
from pydantic import BaseModel
from sqlmodel import Session, select, func
from .deps import get_current_user
from ..database import get_session
from ..models.user import User
from ..models.task import Task
from ..models.send_log import SendLog

router = APIRouter(prefix="/api/dashboard", tags=["仪表盘"])


class StatsResponse(BaseModel):
    today_sent: int
    today_success: int
    today_fail: int
    today_success_rate: float
    week_sent: int
    week_success: int
    week_success_rate: float
    total_tasks: int
    active_tasks: int


class RealtimeResponse(BaseModel):
    running_tasks: int
    queued_tasks: int
    worker_count: int


@router.get("/stats", response_model=StatsResponse)
def get_stats(current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    today_start = datetime.utcnow().replace(hour=0, minute=0, second=0, microsecond=0)
    week_start = today_start - timedelta(days=7)

    # Today's stats
    today_logs = session.exec(
        select(SendLog)
        .where(SendLog.task_id.in_(
            select(Task.id).where(Task.user_id == current_user.id)
        ))
        .where(SendLog.sent_at >= today_start)
    ).all()

    today_sent = len(today_logs)
    today_success = sum(1 for log in today_logs if log.status == "success")
    today_fail = sum(1 for log in today_logs if log.status == "failed")
    today_success_rate = round(today_success / today_sent * 100, 1) if today_sent > 0 else 0

    # Week stats
    week_logs = session.exec(
        select(SendLog)
        .where(SendLog.task_id.in_(
            select(Task.id).where(Task.user_id == current_user.id)
        ))
        .where(SendLog.sent_at >= week_start)
    ).all()

    week_sent = len(week_logs)
    week_success = sum(1 for log in week_logs if log.status == "success")
    week_success_rate = round(week_success / week_sent * 100, 1) if week_sent > 0 else 0

    # Task stats
    total_tasks = session.exec(
        select(func.count(Task.id)).where(Task.user_id == current_user.id)
    ).one()

    active_tasks = session.exec(
        select(func.count(Task.id)).where(Task.user_id == current_user.id).where(Task.status == "running")
    ).one()

    return StatsResponse(
        today_sent=today_sent,
        today_success=today_success,
        today_fail=today_fail,
        today_success_rate=today_success_rate,
        week_sent=week_sent,
        week_success=week_success,
        week_success_rate=week_success_rate,
        total_tasks=total_tasks,
        active_tasks=active_tasks,
    )


@router.get("/realtime", response_model=RealtimeResponse)
def get_realtime(current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    running_tasks = session.exec(
        select(func.count(Task.id)).where(Task.user_id == current_user.id).where(Task.status == "running")
    ).one()

    queued_tasks = session.exec(
        select(func.count(Task.id)).where(Task.user_id == current_user.id).where(Task.status == "pending")
    ).one()

    return RealtimeResponse(running_tasks=running_tasks, queued_tasks=queued_tasks, worker_count=1)