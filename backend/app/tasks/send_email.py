import json
import time
import random
from datetime import datetime
from celery import shared_task
from sqlmodel import Session, select
from ..database import engine
from ..models.task import Task
from ..models.send_log import SendLog
from ..models.sender import Sender
from ..services.load_balancer import LoadBalancer
from ..services.email_sender import email_sender
from ..services.tracker import inject_tracking_pixel, replace_links_with_tracking
from ..websocket.manager import ws_manager


@shared_task(bind=True, max_retries=1)
def send_batch_task(self, task_id: int):
    with Session(engine) as session:
        task = session.get(Task, task_id)
        if not task or task.status not in ("pending", "running"):
            return

        task.status = "running"
        session.add(task)
        session.commit()

        # Load senders
        sender_ids = json.loads(task.sender_ids)
        senders = [session.get(Sender, sid) for sid in sender_ids]
        senders = [s for s in senders if s and s.is_available()]

        if not senders:
            task.status = "completed"
            task.completed_at = datetime.utcnow()
            session.add(task)
            session.commit()
            return

        lb = LoadBalancer(strategy=task.load_balance_strategy)
        email_engine = email_sender

        # Load pending logs for this task
        logs = session.exec(
            select(SendLog).where(SendLog.task_id == task_id).where(SendLog.status == "pending")
        ).all()

        total = len(logs)

        for i, log in enumerate(logs):
            # Check if task is paused or cancelled
            session.refresh(task)
            if task.status in ("paused", "cancelled"):
                return

            # Pick sender via load balancer
            sender = lb.pick_sender(senders)
            if not sender:
                log.status = "failed"
                log.error_message = "No available sender"
                log.sent_at = datetime.utcnow()
                session.add(log)
                session.commit()
                continue

            log.sender_id = sender.id

            # Build personalized body with tracking
            body = task.body
            body = inject_tracking_pixel(body, log.id)
            body = replace_links_with_tracking(body, log.id)

            # Send email
            success, error = email_engine.send(
                sender=sender,
                recipient_email=log.recipient_email,
                recipient_name=log.recipient_name,
                subject=task.subject,
                body_html=body,
                attachments=json.loads(task.attachments) if task.attachments else [],
            )

            log.sent_at = datetime.utcnow()

            if success:
                log.status = "success"
                task.success_count += 1
                lb.report_success(sender, session)
            else:
                log.status = "failed"
                log.error_message = error
                task.fail_count += 1
                lb.report_failure(sender, session)

            session.add(log)
            session.add(task)
            session.commit()

            # WebSocket progress update
            import asyncio
            try:
                loop = asyncio.get_event_loop()
                if loop.is_running():
                    asyncio.ensure_future(ws_manager.send_to_task(task_id, {
                        "type": "progress",
                        "task_id": task_id,
                        "current": i + 1,
                        "total": total,
                        "success": task.success_count,
                        "fail": task.fail_count,
                        "last_email": log.recipient_email,
                        "last_status": log.status,
                    }))
            except RuntimeError:
                pass

            # Delay between sends
            if i < total - 1:
                delay = random.uniform(task.delay_min, task.delay_max)
                time.sleep(delay)

        # Task completed
        task.status = "completed"
        task.completed_at = datetime.utcnow()
        session.add(task)
        session.commit()

        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                asyncio.ensure_future(ws_manager.send_to_task(task_id, {
                    "type": "completed",
                    "task_id": task_id,
                    "success": task.success_count,
                    "fail": task.fail_count,
                }))
        except RuntimeError:
            pass