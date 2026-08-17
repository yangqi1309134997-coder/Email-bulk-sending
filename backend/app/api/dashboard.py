from datetime import timedelta
from fastapi import APIRouter, Depends
from pydantic import BaseModel
from sqlalchemy import case
from sqlmodel import Session, select, func
from .deps import get_current_user
from ..database import get_session
from ..models.user import User
from ..models.task import Task
from ..models.send_log import SendLog
from ..config import settings
from ..utils.time import utcnow

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


class TrendPoint(BaseModel):
    date: str
    sent: int
    success: int
    failed: int


@router.get("/stats", response_model=StatsResponse)
def get_stats(current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    today_start = utcnow().replace(hour=0, minute=0, second=0, microsecond=0)
    week_start = today_start - timedelta(days=7)

    def aggregate(start):
        row = session.exec(
            select(
                func.count(SendLog.id),
                func.coalesce(func.sum(case((SendLog.status == "success", 1), else_=0)), 0),
                func.coalesce(func.sum(case((SendLog.status == "failed", 1), else_=0)), 0),
            )
            .join(Task, Task.id == SendLog.task_id)
            .where(Task.user_id == current_user.id)
            .where(SendLog.sent_at >= start)
        ).one()
        return int(row[0] or 0), int(row[1] or 0), int(row[2] or 0)

    today_sent, today_success, today_fail = aggregate(today_start)
    today_success_rate = round(today_success / today_sent * 100, 1) if today_sent > 0 else 0

    week_sent, week_success, _week_fail = aggregate(week_start)
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


@router.get("/trend", response_model=list[TrendPoint])
def get_trend(current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    """Daily send volume for the last 7 days (inclusive of today)."""
    today_start = utcnow().replace(hour=0, minute=0, second=0, microsecond=0)
    week_start = today_start - timedelta(days=6)
    rows = session.exec(
        select(SendLog.sent_at, SendLog.status)
        .join(Task, Task.id == SendLog.task_id)
        .where(Task.user_id == current_user.id)
        .where(SendLog.sent_at >= week_start)
    ).all()
    buckets = {
        (week_start + timedelta(days=i)).date().isoformat(): {"sent": 0, "success": 0, "failed": 0}
        for i in range(7)
    }
    for sent_at, status in rows:
        if not sent_at:
            continue
        key = sent_at.date().isoformat()
        if key not in buckets:
            continue
        buckets[key]["sent"] += 1
        if status == "success":
            buckets[key]["success"] += 1
        elif status == "failed":
            buckets[key]["failed"] += 1
    return [TrendPoint(date=day, **counts) for day, counts in buckets.items()]


@router.get("/realtime", response_model=RealtimeResponse)
def get_realtime(current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    running_tasks = session.exec(
        select(func.count(Task.id)).where(Task.user_id == current_user.id).where(Task.status == "running")
    ).one()

    queued_tasks = session.exec(
        select(func.count(Task.id)).where(Task.user_id == current_user.id).where(Task.status == "pending")
    ).one()

    return RealtimeResponse(
        running_tasks=running_tasks,
        queued_tasks=queued_tasks,
        worker_count=max(1, min(32, int(getattr(settings, "SEND_ENGINE_WORKERS", 4) or 4))),
    )
