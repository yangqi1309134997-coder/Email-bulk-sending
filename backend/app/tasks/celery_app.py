from celery import Celery
from celery.schedules import crontab
from ..config import settings

celery_app = Celery(
    "email_bulk_sending",
    broker=settings.CELERY_BROKER_URL,
    backend=settings.CELERY_RESULT_BACKEND,
    include=["app.tasks.send_email", "app.tasks.scheduled_tasks"],
)

celery_app.conf.update(
    task_serializer="json",
    accept_content=["json"],
    result_serializer="json",
    timezone="UTC",
    enable_utc=True,
    task_track_started=True,
    task_acks_late=True,
    task_reject_on_worker_lost=True,
    task_acks_on_failure_or_timeout=False,
    worker_prefetch_multiplier=1,
    task_time_limit=None,
    task_soft_time_limit=None,
)

# Celery Beat Schedule for background auto-send
celery_app.conf.beat_schedule = {
    # Check for scheduled tasks every minute
    "check-scheduled-tasks": {
        "task": "app.tasks.scheduled_tasks.check_scheduled_tasks",
        "schedule": 60.0,  # every 60 seconds
    },
    # Clean up old send logs daily at 3 AM
    "cleanup-old-logs": {
        "task": "app.tasks.scheduled_tasks.cleanup_old_logs",
        "schedule": crontab(hour=3, minute=0),
    },
    # Reset daily quotas at midnight
    "reset-daily-quotas": {
        "task": "app.tasks.scheduled_tasks.reset_daily_quotas",
        "schedule": crontab(hour=0, minute=0),
    },
    # Recover stuck tasks every 5 minutes
    "recover-stuck-tasks": {
        "task": "app.tasks.scheduled_tasks.recover_stuck_tasks",
        "schedule": 300.0,  # every 5 minutes
    },
}
