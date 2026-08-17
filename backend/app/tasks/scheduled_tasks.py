from celery import shared_task
import json

from sqlalchemy import delete, func, update
from sqlmodel import Session, select
from datetime import datetime, timedelta
from ..database import engine
from ..models.task import Task
from ..models.sender import Sender
from ..models.send_log import SendLog
from ..services.send_engine import _coerce_bool, _parse_smart_config, get_send_engine
from ..utils.time import utcnow

logger = None

def get_logger():
    global logger
    if logger is None:
        import logging
        logger = logging.getLogger(__name__)
    return logger


def parse_sender_ids(raw) -> list[int]:
    try:
        values = json.loads(raw or "[]")
    except (TypeError, json.JSONDecodeError):
        return []
    if not isinstance(values, list):
        return []

    sender_ids = []
    seen = set()
    for value in values:
        # bool is an int subclass and floats are truncated by int(); neither
        # is a valid persisted primary key representation.
        if isinstance(value, bool):
            continue
        if isinstance(value, int):
            sender_id = value
        elif isinstance(value, str):
            normalized = value.strip()
            if not normalized.isascii() or not normalized.isdecimal():
                continue
            try:
                sender_id = int(normalized)
            except ValueError:
                continue
        else:
            continue
        if sender_id > 0 and sender_id not in seen:
            seen.add(sender_id)
            sender_ids.append(sender_id)
    return sender_ids


def should_auto_resume_task(task, now: datetime | None = None) -> bool:
    if getattr(task, "pause_reason", "") == "manual":
        return False
    smart = _parse_smart_config(getattr(task, "smart_config", "{}"))
    if not _coerce_bool(smart.get("auto_resume_after_cooldown", True), True):
        return False
    run_at = getattr(task, "next_run_at", None)
    return not run_at or run_at <= (now or utcnow())


@shared_task(name="app.tasks.scheduled_tasks.check_scheduled_tasks")
def check_scheduled_tasks():
    """Check for tasks that are scheduled to run now and auto-resume risk-paused tasks."""
    log = get_logger()
    with Session(engine) as session:
        now = utcnow()

        # 1) Scheduled tasks due now
        tasks = session.exec(
            select(Task).where(
                Task.status == "pending",
                Task.schedule_type == "scheduled",
                Task.schedule_time <= now,
            ).order_by(Task.schedule_time).limit(500)
        ).all()

        for task in tasks:
            sender_ids = parse_sender_ids(task.sender_ids)
            if not sender_ids:
                log.warning("Task %s has invalid sender_ids JSON", task.id)

            senders = [session.get(Sender, sid) for sid in sender_ids]
            # Auto-unpause expired senders before availability check
            for s in senders:
                if s and s.status == "paused" and s.paused_until and s.paused_until <= now:
                    s.status = "active"
                    s.paused_until = None
                    s.consecutive_failures = 0
                    session.add(s)
            session.commit()

            available = [s for s in senders if s and s.is_available()]
            if not available:
                # Keep pending if schedule just hit but all senders cooling down — retry later
                log.warning("Task %s has no available senders at schedule time, will retry", task.id)
                continue

            get_send_engine().submit(task.id)
            log.info("Auto-started scheduled task %s", task.id)

        # 2) Smart schedule type: treat like immediate when pending and no future schedule_time
        smart_tasks = session.exec(
            select(Task).where(
                Task.status == "pending",
                Task.schedule_type == "smart",
            ).order_by(Task.created_at).limit(500)
        ).all()
        for task in smart_tasks:
            # Start smart tasks immediately (domain/timezone optimization can be layered later)
            get_send_engine().submit(task.id)
            log.info("Auto-started smart task %s", task.id)

        # 3) Risk-paused tasks with pending logs + auto_resume
        paused = session.exec(
            select(Task).where(Task.status == "paused").order_by(Task.id).limit(500)
        ).all()
        for task in paused:
            if not should_auto_resume_task(task, now):
                continue
            pending = session.exec(
                select(SendLog)
                .where(SendLog.task_id == task.id)
                .where(SendLog.status == "pending")
                .limit(1)
            ).first()
            if pending:
                # SendEngine itself enforces resume delay if still scheduled
                get_send_engine().submit(task.id)
                log.info("Re-queued paused task %s for auto-resume check", task.id)


@shared_task(name="app.tasks.scheduled_tasks.cleanup_old_logs")
def cleanup_old_logs():
    """Clean up old send logs (older than 90 days)."""
    log = get_logger()
    cutoff = utcnow() - timedelta(days=90)
    with Session(engine) as session:
        result = session.exec(
            delete(SendLog).where(
                SendLog.sent_at != None,  # noqa: E711
                SendLog.sent_at < cutoff,
            )
        )
        count = max(0, int(result.rowcount or 0))
        session.commit()
        log.info("Cleaned up %s old send logs", count)
    return {"cleaned": count}


@shared_task(name="app.tasks.scheduled_tasks.reset_daily_quotas")
def reset_daily_quotas():
    """Reset daily sent counts for all senders and clear expired pauses."""
    log = get_logger()
    with Session(engine) as session:
        now = utcnow()
        count = int(session.exec(select(func.count(Sender.id))).one() or 0)
        session.exec(update(Sender).values(daily_sent=0, consecutive_failures=0))
        session.exec(
            update(Sender)
            .where(
                Sender.status == "paused",
                Sender.paused_until != None,  # noqa: E711
                Sender.paused_until < now,
            )
            .values(status="active", paused_until=None)
        )
        session.exec(
            update(Sender)
            .where(
                Sender.cb_state == "open",
                Sender.cb_next_attempt_time != None,  # noqa: E711
                Sender.cb_next_attempt_time < now,
            )
            .values(
                cb_state="closed",
                cb_failure_count=0,
                cb_success_count=0,
                cb_next_attempt_time=None,
            )
        )
        session.commit()
        log.info("Reset daily quotas for %s senders", count)
    return {"reset": count}


@shared_task(name="app.tasks.scheduled_tasks.recover_stuck_tasks")
def recover_stuck_tasks():
    """Recover tasks that are stuck in running/paused state with pending logs."""
    log = get_logger()
    se = get_send_engine()
    with Session(engine) as session:
        recovered = 0
        tasks = session.exec(
            select(Task)
            .where(Task.status.in_(["running", "paused"]))  # type: ignore[attr-defined]
            .order_by(Task.id)
            .limit(1000)
        ).all()
        for task in tasks:
            if task.status == "paused" and not should_auto_resume_task(task):
                continue
            pending = session.exec(
                select(SendLog)
                .where(SendLog.task_id == task.id)
                .where(SendLog.status == "pending")
                .limit(1)
            ).first()
            if pending:
                se.submit(task.id)
                recovered += 1
                log.info("Recovered stuck task %s (status=%s)", task.id, task.status)
        return {"recovered": recovered}
