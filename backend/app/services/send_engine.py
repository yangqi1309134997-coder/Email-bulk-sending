"""Background email sending engine with concurrency, risk control, and crash recovery."""

from __future__ import annotations

import json
import logging
import random
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from types import SimpleNamespace
from typing import Any, Optional

from sqlalchemy import func, or_, text, update
from sqlmodel import Session, select

from ..config import settings
from ..database import engine
from ..models.sender import Sender
from ..models.send_log import SendLog
from ..models.task import Task
from ..services.email_sender import email_sender, is_auth_error, is_rate_limit_error
from ..services.load_balancer import CircuitBreakerState, LoadBalancer
from ..services.tracker import inject_tracking_pixel, replace_links_with_tracking
from ..websocket.events import publish_task_event
from ..utils.time import from_unix_utc, to_unix_utc, utcnow

logger = logging.getLogger(__name__)

_engine_instance: Optional["SendEngine"] = None
_engine_lock = threading.Lock()


def _parse_smart_config(raw: Any) -> dict:
    if not raw:
        return {}
    if isinstance(raw, dict):
        return raw
    try:
        parsed = json.loads(raw)
        return parsed if isinstance(parsed, dict) else {}
    except Exception:
        return {}


def _bounded_int(value: Any, default: int, minimum: int, maximum: int) -> int:
    try:
        return max(minimum, min(maximum, int(value)))
    except (TypeError, ValueError):
        return default


def _bounded_float(value: Any, default: float, minimum: float, maximum: float) -> float:
    try:
        return max(minimum, min(maximum, float(value)))
    except (TypeError, ValueError):
        return default


def _coerce_bool(value: Any, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"true", "1", "yes", "on"}:
            return True
        if normalized in {"false", "0", "no", "off"}:
            return False
    return default


def _utcnow() -> datetime:
    return utcnow()

def _recipient_timezone_offset(email: str) -> int:
    """Rough timezone offset hours from common email domains. Best-effort only."""
    domain = (email or "").split("@")[-1].lower()

    def matches(*suffixes: str) -> bool:
        return any(domain == suffix or domain.endswith("." + suffix) for suffix in suffixes)

    # China / Asia
    if domain.endswith((".cn", ".com.cn")) or matches(
        "qq.com", "163.com", "126.com", "yeah.net", "sina.com", "aliyun.com", "foxmail.com", "139.com", "189.cn"
    ):
        return 8
    if domain.endswith((".jp", ".co.jp")):
        return 9
    if domain.endswith((".kr", ".co.kr")):
        return 9
    if domain.endswith((".in", ".co.in")):
        return 5
    if domain.endswith((".sg", ".com.sg")):
        return 8
    if domain.endswith((".au", ".com.au")):
        return 10
    # Europe
    if domain.endswith((".uk", ".de", ".fr", ".it", ".es", ".nl", ".eu")):
        return 1
    # Americas
    if domain.endswith(".us") or matches(
        "gmail.com", "outlook.com", "hotmail.com", "yahoo.com", "icloud.com", "aol.com"
    ):
        return -5
    if domain.endswith((".br", ".com.br")):
        return -3
    return 8


def _is_good_send_hour(email: str, now: Optional[datetime] = None) -> bool:
    """Prefer local 9-11 and 14-17 windows for smart schedule."""
    now = now or _utcnow()
    offset = _recipient_timezone_offset(email)
    local_hour = (now.hour + offset) % 24
    return (9 <= local_hour < 12) or (14 <= local_hour < 18)


def _pick_proxy(proxies: list[str], index: int) -> str:
    if not proxies:
        return ""
    return proxies[index % len(proxies)]




class SendEngine:
    """Persistent background engine that processes tasks from an in-process queue."""

    def __init__(self, max_workers: int = 4):
        self._queue: list[int] = []
        self._lock = threading.RLock()
        self._active_tasks: set[int] = set()
        self._resume_at: dict[int, float] = {}  # task_id -> unix ts
        self._running = True
        self._stop_event = threading.Event()
        self._send_slots = threading.BoundedSemaphore(
            max(1, int(settings.MAX_GLOBAL_SEND_CONCURRENCY))
        )
        self._sender_slot_limit = max(1, int(settings.SMTP_POOL_MAX_PER_SENDER))
        self._sender_slots: dict[int, threading.BoundedSemaphore] = {}
        self._sender_rate_locks: dict[int, threading.Lock] = {}
        self._sender_next_send_at: dict[int, float] = {}
        self._engine_id = f"{uuid.uuid4().hex}:{threading.get_native_id()}"
        self._workers: list[threading.Thread] = []
        for i in range(max_workers):
            t = threading.Thread(target=self._worker_loop, name=f"send-worker-{i}", daemon=True)
            t.start()
            self._workers.append(t)
        self._recovery_thread = threading.Thread(
            target=self._recovery_loop,
            name="send-recovery",
            daemon=True,
        )
        self._recovery_thread.start()
        logger.info("SendEngine started with %s workers", max_workers)

    def _sender_runtime(
        self, sender_id: int
    ) -> tuple[threading.BoundedSemaphore, threading.Lock]:
        """Return process-wide concurrency and pacing controls for a sender."""
        with self._lock:
            slots = getattr(self, "_sender_slots", None)
            if slots is None:
                slots = {}
                self._sender_slots = slots
            rate_locks = getattr(self, "_sender_rate_locks", None)
            if rate_locks is None:
                rate_locks = {}
                self._sender_rate_locks = rate_locks
            slot = slots.get(sender_id)
            if slot is None:
                limit = max(
                    1,
                    int(
                        getattr(
                            self,
                            "_sender_slot_limit",
                            settings.SMTP_POOL_MAX_PER_SENDER,
                        )
                    ),
                )
                slot = threading.BoundedSemaphore(limit)
                slots[sender_id] = slot
            rate_lock = rate_locks.get(sender_id)
            if rate_lock is None:
                rate_lock = threading.Lock()
                rate_locks[sender_id] = rate_lock
            if not hasattr(self, "_sender_next_send_at"):
                self._sender_next_send_at = {}
            return slot, rate_lock

    def submit(self, task_id: int, delay_seconds: float = 0) -> None:
        with self._lock:
            if delay_seconds > 0:
                self._resume_at[task_id] = time.time() + delay_seconds
            else:
                self._resume_at.pop(task_id, None)
            if task_id not in self._queue:
                self._queue.append(task_id)
                logger.info("Task %s submitted (delay=%ss)", task_id, delay_seconds)
            elif task_id in self._queue and delay_seconds <= 0:
                self._resume_at.pop(task_id, None)

    def _pop_task(self) -> Optional[int]:
        now = time.time()
        with self._lock:
            for i, tid in enumerate(list(self._queue)):
                if tid in self._active_tasks:
                    continue
                resume_at = self._resume_at.get(tid)
                if resume_at and resume_at > now:
                    continue
                self._queue.pop(i)
                self._resume_at.pop(tid, None)
                self._active_tasks.add(tid)
                return tid
        return None

    def _mark_done(self, task_id: int) -> None:
        with self._lock:
            self._active_tasks.discard(task_id)

    def _try_acquire_task_lease(self, task_id: int) -> bool:
        from ..config import settings

        now = _utcnow()
        expires_at = now + timedelta(seconds=max(30, int(settings.SEND_TASK_LEASE_SECONDS)))
        with engine.begin() as connection:
            result = connection.execute(
                text(
                    "UPDATE send_tasks SET lease_owner=:owner, lease_expires_at=:expires, "
                    "last_heartbeat_at=:now WHERE id=:task_id AND ("
                    "lease_owner IS NULL OR lease_owner='' OR lease_owner=:owner OR "
                    "lease_expires_at IS NULL OR lease_expires_at<=:now)"
                ),
                {
                    "owner": self._engine_id,
                    "expires": expires_at,
                    "now": now,
                    "task_id": task_id,
                },
            )
            return result.rowcount == 1

    def _heartbeat_task_lease(self, task_id: int) -> bool:
        from ..config import settings

        now = _utcnow()
        expires_at = now + timedelta(seconds=max(30, int(settings.SEND_TASK_LEASE_SECONDS)))
        with engine.begin() as connection:
            result = connection.execute(
                text(
                    "UPDATE send_tasks SET lease_expires_at=:expires, last_heartbeat_at=:now "
                    "WHERE id=:task_id AND lease_owner=:owner"
                ),
                {
                    "owner": self._engine_id,
                    "expires": expires_at,
                    "now": now,
                    "task_id": task_id,
                },
            )
            return result.rowcount == 1

    def _lease_heartbeat_interval(self) -> float:
        lease_seconds = max(30, int(settings.SEND_TASK_LEASE_SECONDS))
        return max(2.0, min(10.0, lease_seconds / 3.0))

    def _lease_heartbeat_loop(self, task_id: int, stop_event: threading.Event) -> None:
        """Keep a task lease valid while a network batch is still in flight."""
        interval = self._lease_heartbeat_interval()
        while not stop_event.wait(interval):
            try:
                if not self._heartbeat_task_lease(task_id):
                    logger.error("Task %s lease ownership was lost", task_id)
                    return
            except Exception:
                logger.exception("Task %s lease heartbeat failed", task_id)

    def _release_task_lease(self, task_id: int) -> None:
        with engine.begin() as connection:
            connection.execute(
                text(
                    "UPDATE send_tasks SET lease_owner='', lease_expires_at=NULL "
                    "WHERE id=:task_id AND lease_owner=:owner"
                ),
                {"owner": self._engine_id, "task_id": task_id},
            )

    def _worker_loop(self) -> None:
        while self._running:
            task_id = self._pop_task()
            if task_id is None:
                self._stop_event.wait(0.5)
                continue
            try:
                self._process_task(task_id)
            except Exception:
                logger.exception("Task %s crashed", task_id)
                try:
                    with Session(engine) as session:
                        task = session.get(Task, task_id)
                        if task and task.status == "running":
                            # Keep recoverable if pending logs remain
                            pending = session.exec(
                                select(SendLog)
                                .where(SendLog.task_id == task_id)
                                .where(SendLog.status == "pending")
                                .limit(1)
                            ).first()
                            if pending:
                                task.status = "paused"
                                if hasattr(task, "pause_reason"):
                                    task.pause_reason = "engine_crash"
                                session.add(task)
                                session.commit()
                                self.submit(task_id, delay_seconds=5)
                            else:
                                task.status = "failed"
                                task.completed_at = _utcnow()
                                session.add(task)
                                session.commit()
                except Exception:
                    logger.exception("Failed to mark task %s failed", task_id)
            finally:
                self._mark_done(task_id)

    def _refresh_senders(self, session: Session, sender_ids: list[int]) -> list[Sender]:
        senders = []
        changed = False
        for sid in sender_ids:
            s = session.get(Sender, sid)
            if not s:
                continue
            if s.status == "paused" and s.paused_until and s.paused_until <= _utcnow():
                s.status = "active"
                s.paused_until = None
                s.consecutive_failures = 0
                session.add(s)
                changed = True
            if s.is_available():
                senders.append(s)
        if changed:
            session.commit()
        return senders

    def _snapshot_sender(self, sender: Sender) -> SimpleNamespace:
        """Detach sender fields for thread-safe use outside the ORM session."""
        return SimpleNamespace(
            id=sender.id,
            email=sender.email,
            password=sender.password,
            smtp_server=sender.smtp_server,
            smtp_port=sender.smtp_port,
            use_tls=sender.use_tls,
            sender_type=sender.sender_type,
            enabled=sender.enabled,
            weight=sender.weight,
            daily_quota=sender.daily_quota,
            daily_sent=sender.daily_sent,
            success_rate=sender.success_rate,
            status=sender.status,
            consecutive_failures=sender.consecutive_failures,
            paused_until=sender.paused_until,
            aliyun_access_key=getattr(sender, "aliyun_access_key", "") or "",
            aliyun_access_secret=getattr(sender, "aliyun_access_secret", "") or "",
            aliyun_region=getattr(sender, "aliyun_region", "cn-hangzhou") or "cn-hangzhou",
            aliyun_from_name=getattr(sender, "aliyun_from_name", "") or "",
            smtp_username=getattr(sender, "smtp_username", "") or "",
            smtp_security=getattr(sender, "smtp_security", "") or "",
        )

    def _reserve_sender_capacity(self, session: Session, sender: Sender) -> bool:
        """Atomically reserve one daily attempt before any network send."""
        result = session.exec(
            update(Sender)
            .where(Sender.id == sender.id)
            .where(Sender.enabled == True)  # noqa: E712
            .where(Sender.status.in_(["active", ""]))  # type: ignore[attr-defined]
            .where(
                or_(
                    Sender.daily_quota <= 0,
                    func.coalesce(Sender.daily_sent, 0) < Sender.daily_quota,
                )
            )
            .values(daily_sent=func.coalesce(Sender.daily_sent, 0) + 1)
            .execution_options(synchronize_session=False)
        )
        if result.rowcount != 1:
            return False
        session.flush()
        session.refresh(sender)
        return True

    def _recover_abandoned_logs(self, session: Session, task_id: int) -> None:
        """Do not automatically resend SMTP deliveries with an unknown outcome."""
        abandoned = session.exec(
            select(SendLog).where(
                SendLog.task_id == task_id,
                SendLog.status == "processing",
            )
        ).all()
        if not abandoned:
            return
        now = _utcnow()
        for log in abandoned:
            log.status = "failed"
            log.error_message = "发送进程中断，投递结果未知；请人工确认后重试"
            log.last_error_code = "delivery_outcome_unknown"
            log.sent_at = now
            log.claimed_at = None
            session.add(log)
        session.commit()

    def _pause_task(
        self,
        task: Task,
        session: Session,
        *,
        reason: str,
        delay_seconds: float,
        auto_resume: bool,
    ) -> None:
        delay_seconds = max(0.0, float(delay_seconds))
        task.status = "paused"
        task.pause_reason = reason
        task.next_run_at = _utcnow() + timedelta(seconds=delay_seconds) if auto_resume else None
        session.add(task)
        session.commit()
        if auto_resume:
            self.submit(task.id, delay_seconds=delay_seconds)

    def _send_one(
        self,
        *,
        log_id: int,
        sender_snapshot: SimpleNamespace,
        subject: str,
        body_html: str,
        recipient_email: str,
        recipient_name: str,
        attachments: list,
        max_retries: int,
        retry_backoff_base: float,
        proxy_url: str = "",
    ) -> tuple[int, bool, str, int]:
        """Send a single email using detached snapshots. Thread-safe."""
        body = body_html or ""
        body = inject_tracking_pixel(body, log_id)
        body = replace_links_with_tracking(body, log_id)

        try:
            if sender_snapshot.sender_type in ("阿里云邮箱推送", "aliyun_dm", "Aliyun DM"):
                if attachments:
                    return (
                        log_id,
                        False,
                        "Aliyun DirectMail API does not support attachments; use its SMTP preset",
                        int(sender_snapshot.id),
                    )
                from ..services.aliyun_dm import aliyun_dm_sender

                success, error = aliyun_dm_sender.send(
                    sender=sender_snapshot,
                    recipient_email=recipient_email,
                    recipient_name=recipient_name,
                    subject=subject,
                    body_html=body,
                    max_retries=0,
                    retry_backoff_base=retry_backoff_base,
                )
            else:
                success, error = email_sender.send(
                    sender=sender_snapshot,
                    recipient_email=recipient_email,
                    recipient_name=recipient_name,
                    subject=subject,
                    body_html=body,
                    attachments=attachments,
                    max_retries=max_retries,
                    retry_backoff_base=retry_backoff_base,
                    proxy_url=proxy_url,
                )
        except Exception as e:
            success, error = False, str(e)
        return log_id, success, error or "", int(sender_snapshot.id)

    def _process_task(self, task_id: int) -> None:
        if not self._try_acquire_task_lease(task_id):
            logger.info("Task %s is leased by another engine", task_id)
            return
        heartbeat_stop = threading.Event()
        heartbeat_thread = threading.Thread(
            target=self._lease_heartbeat_loop,
            args=(task_id, heartbeat_stop),
            name=f"task-lease-{task_id}",
            daemon=True,
        )
        heartbeat_thread.start()
        try:
            self._process_task_leased(task_id)
        finally:
            heartbeat_stop.set()
            heartbeat_thread.join(timeout=1.0)
            self._release_task_lease(task_id)

    def _process_task_leased(self, task_id: int) -> None:
        with Session(engine) as session:
            task = session.get(Task, task_id)
            if not task or task.status not in ("pending", "running", "paused"):
                return

            smart = _parse_smart_config(task.smart_config)
            if task.status == "paused":
                auto_resume = _coerce_bool(
                    smart.get("auto_resume_after_cooldown", True), True
                )
                if not auto_resume:
                    return
                if task.next_run_at and task.next_run_at > _utcnow():
                    self.submit(
                        task_id,
                        delay_seconds=max(0.1, (task.next_run_at - _utcnow()).total_seconds()),
                    )
                    return
                with self._lock:
                    resume_at = self._resume_at.get(task_id)
                if resume_at and resume_at > time.time():
                    # Still cooling down; requeue for later
                    self.submit(task_id, delay_seconds=max(1.0, resume_at - time.time()))
                    return
                task.status = "running"
                task.pause_reason = ""
                task.next_run_at = None
                session.add(task)
                session.commit()

            task.status = "running"
            task.last_heartbeat_at = _utcnow()
            session.add(task)
            session.commit()
            self._recover_abandoned_logs(session, task_id)

            try:
                raw_sender_ids = json.loads(task.sender_ids or "[]")
                sender_ids = [
                    int(sender_id)
                    for sender_id in raw_sender_ids
                    if int(sender_id) > 0
                ] if isinstance(raw_sender_ids, list) else []
            except (TypeError, ValueError, json.JSONDecodeError):
                sender_ids = []

            senders = self._refresh_senders(session, sender_ids)
            if not senders:
                pending = session.exec(
                    select(SendLog).where(SendLog.task_id == task_id, SendLog.status == "pending").limit(1)
                ).first()
                if pending:
                    wait_seconds = _bounded_int(
                        smart.get("risk_pause_seconds", 300), 300, 30, 7200
                    )
                    self._pause_task(
                        task,
                        session,
                        reason="no_available_sender",
                        delay_seconds=wait_seconds,
                        auto_resume=_coerce_bool(
                            smart.get("auto_resume_after_cooldown", True), True
                        ),
                    )
                    logger.warning("Task %s: no senders available, auto-pause then retry", task_id)
                    return
                task.status = "completed"
                task.completed_at = _utcnow()
                session.add(task)
                session.commit()
                return

            lb = LoadBalancer(strategy=task.load_balance_strategy or "round_robin")
            self._restore_circuit_breakers(lb, senders, session)

            max_retries = _bounded_int(smart.get("max_retries", 3), 3, 0, 10)
            retry_backoff_base = _bounded_float(
                smart.get("retry_backoff_base", 2), 2.0, 1.0, 60.0
            )
            rate_limit_cooldown = _bounded_int(
                smart.get("rate_limit_cooldown", 60), 60, 10, 3600
            )
            max_consecutive_rate_limits = _bounded_int(
                smart.get("max_consecutive_rate_limits", 5), 5, 1, 20
            )
            auto_resume_after_cooldown = _coerce_bool(
                smart.get("auto_resume_after_cooldown", True), True
            )
            risk_auto_pause_task = _coerce_bool(
                smart.get("risk_auto_pause_task", True), True
            )
            risk_pause_seconds = _bounded_int(
                smart.get("risk_pause_seconds", 300), 300, 30, 7200
            )
            concurrency = _bounded_int(
                smart.get("concurrency_per_sender", 1), 1, 1, 20
            )
            concurrency = min(
                concurrency,
                max(1, int(settings.SMTP_POOL_MAX_PER_SENDER)),
            )
            batch_size = _bounded_int(smart.get("batch_size", 100), 100, 1, 500)
            rate_limit_patterns = smart.get(
                "rate_limit_patterns",
                [
                    "Too many attempts",
                    "rate limit",
                    "spam",
                    "blocked",
                    "too many",
                    "quota",
                    "limit exceeded",
                ],
            )
            if not isinstance(rate_limit_patterns, list):
                rate_limit_patterns = []
            rate_limit_patterns = [
                str(pattern).strip()[:200]
                for pattern in rate_limit_patterns
                if str(pattern).strip()
            ][:50]

            try:
                raw_attachments = json.loads(task.attachments) if task.attachments else []
                if not isinstance(raw_attachments, list):
                    raw_attachments = []
            except Exception:
                raw_attachments = []

            # Tasks created by older versions may contain arbitrary absolute
            # paths. Re-check ownership immediately before sending so a
            # database edit or a moved symlink can never exfiltrate a server
            # file. Missing files are skipped and the task remains sendable.
            attachments: list[str] = []
            try:
                from ..api.upload import resolve_user_attachment_paths

                for supplied_path in raw_attachments:
                    try:
                        attachments.extend(
                            resolve_user_attachment_paths(task.user_id, [str(supplied_path)])
                        )
                    except ValueError:
                        logger.warning(
                            "Skipping unsafe or missing attachment for task %s: %s",
                            task_id,
                            supplied_path,
                        )
            except Exception:
                logger.exception("Unable to validate task %s attachments", task_id)
                attachments = []
            attachments = list(dict.fromkeys(attachments))

            consecutive_rate_limits = 0
            subject = task.subject or ""
            body_template = task.body or ""
            delay_min = max(0, int(task.delay_min or 0))
            delay_max = max(delay_min, int(task.delay_max or 0))
            try:
                proxies = json.loads(task.proxies or "[]")
                if not isinstance(proxies, list):
                    proxies = []
            except Exception:
                proxies = []
            proxies = [str(p).strip() for p in proxies if str(p).strip()]
            schedule_type = task.schedule_type or "immediate"
            publish_task_event(
                task_id,
                {
                    "type": "status",
                    "status": "running",
                    "success_count": task.success_count,
                    "fail_count": task.fail_count,
                    "recipient_count": task.recipient_count,
                    "message": "任务开始发送",
                },
            )

            while True:
                session.refresh(task)
                task.last_heartbeat_at = _utcnow()
                session.add(task)
                session.commit()
                self._heartbeat_task_lease(task_id)

                if task.status in ("cancelled", "completed"):
                    self._persist_circuit_breakers(lb, session)
                    return
                if task.status == "paused":
                    self._persist_circuit_breakers(lb, session)
                    return

                now = _utcnow()
                due_statement = (
                    select(SendLog)
                    .where(SendLog.task_id == task_id)
                    .where(SendLog.status == "pending")
                    .where(
                        or_(
                            SendLog.next_attempt_at == None,  # noqa: E711
                            SendLog.next_attempt_at <= now,
                        )
                    )
                    .order_by(SendLog.id)
                )
                if schedule_type == "smart":
                    # Domain-based time-window selection happens in Python.
                    # Stream the complete due set so early rows from one time
                    # zone cannot starve eligible recipients later in a large
                    # task, while keeping memory bounded.
                    logs = []
                    due_exists = False
                    due_result = session.exec(
                        due_statement.execution_options(yield_per=1000)
                    )
                    for candidate in due_result:
                        due_exists = True
                        if _is_good_send_hour(candidate.recipient_email, now):
                            logs.append(candidate)
                            if len(logs) >= batch_size:
                                break
                else:
                    logs = session.exec(due_statement.limit(batch_size)).all()
                    due_exists = bool(logs)

                if not due_exists:
                    next_attempt = session.exec(
                        select(func.min(SendLog.next_attempt_at)).where(
                            SendLog.task_id == task_id,
                            SendLog.status == "pending",
                            SendLog.next_attempt_at != None,  # noqa: E711
                        )
                    ).one()
                    if next_attempt:
                        wait_seconds = max(0.1, (next_attempt - _utcnow()).total_seconds())
                        self._pause_task(
                            task,
                            session,
                            reason="retry_backoff",
                            delay_seconds=wait_seconds,
                            auto_resume=auto_resume_after_cooldown,
                        )
                        self._persist_circuit_breakers(lb, session)
                        return
                    break

                # Smart schedule: prefer recipients currently in good local hours
                if schedule_type == "smart":
                    if not logs:
                        # No recipient in good window now; wait and retry later
                        wait_s = 15 * 60
                        self._pause_task(
                            task,
                            session,
                            reason="smart_window_wait",
                            delay_seconds=wait_s,
                            auto_resume=auto_resume_after_cooldown,
                        )
                        self._persist_circuit_breakers(lb, session)
                        publish_task_event(
                            task_id,
                            {
                                "type": "status",
                                "status": "paused",
                                "message": f"智能分时段等待更佳发送窗口({wait_s}s)",
                                "success_count": task.success_count,
                                "fail_count": task.fail_count,
                                "recipient_count": task.recipient_count,
                            },
                        )
                        return

                senders = self._refresh_senders(session, sender_ids)
                if not senders:
                    if risk_auto_pause_task and auto_resume_after_cooldown:
                        self._pause_task(
                            task,
                            session,
                            reason="no_available_sender",
                            delay_seconds=risk_pause_seconds,
                            auto_resume=True,
                        )
                        self._persist_circuit_breakers(lb, session)
                        logger.warning(
                            "Task %s paused: no available senders, resume in %ss",
                            task_id,
                            risk_pause_seconds,
                        )
                        return
                    break

                worker_limit = min(
                    batch_size,
                    max(1, int(settings.MAX_TASK_CONCURRENCY)),
                    max(1, concurrency * len(senders)),
                )
                logs = logs[:worker_limit]

                # Snapshot assignments with plain data so worker threads never touch ORM instances
                assignments: list[dict] = []
                for log in logs:
                    session.refresh(task)
                    if task.status in ("paused", "cancelled"):
                        self._persist_circuit_breakers(lb, session)
                        return
                    candidates = list(senders)
                    sender = None
                    while candidates:
                        candidate = lb.pick_sender(candidates)
                        if candidate is None:
                            break
                        if self._reserve_sender_capacity(session, candidate):
                            sender = candidate
                            break
                        candidates = [item for item in candidates if item.id != candidate.id]
                        senders = [item for item in senders if item.id != candidate.id]
                    if not sender:
                        break
                    log.sender_id = sender.id
                    log.attempt_count = int(log.attempt_count or 0) + 1
                    log.status = "processing"
                    log.claimed_at = _utcnow()
                    log.next_attempt_at = None
                    session.add(log)
                    assignments.append(
                        {
                            "log_id": log.id,
                            "recipient_email": log.recipient_email,
                            "recipient_name": log.recipient_name or "",
                            "sender": self._snapshot_sender(sender),
                            "proxy_url": _pick_proxy(proxies, len(assignments)),
                        }
                    )
                session.commit()

                if not assignments:
                    self._pause_task(
                        task,
                        session,
                        reason="no_available_sender",
                        delay_seconds=risk_pause_seconds,
                        auto_resume=auto_resume_after_cooldown,
                    )
                    self._persist_circuit_breakers(lb, session)
                    return

                session.refresh(task)
                if task.status in ("paused", "cancelled"):
                    self._persist_circuit_breakers(lb, session)
                    return

                per_sender_limits = {
                    sender.id: threading.BoundedSemaphore(concurrency) for sender in senders
                }

                def execute_assignment(item):
                    sender_id = item["sender"].id
                    semaphore = per_sender_limits.setdefault(
                        sender_id,
                        threading.BoundedSemaphore(concurrency),
                    )
                    sender_slots, rate_lock = self._sender_runtime(sender_id)
                    with semaphore:
                        sender_slots.acquire()
                        try:
                            if delay_max > 0:
                                with rate_lock:
                                    current = time.monotonic()
                                    next_send_at = self._sender_next_send_at.get(sender_id, 0.0)
                                    wait_for = max(0.0, next_send_at - current)
                                    # Delay jitter is operational, not security-sensitive.
                                    spacing = random.uniform(delay_min, delay_max)  # nosec B311
                                    self._sender_next_send_at[sender_id] = max(
                                        current, next_send_at
                                    ) + spacing
                                if wait_for > 0:
                                    stop_event = getattr(self, "_stop_event", None)
                                    if stop_event is not None:
                                        stop_event.wait(wait_for)
                                    else:
                                        time.sleep(wait_for)

                            # Delayed work must not occupy the process-wide
                            # network semaphore and block unrelated tasks.
                            send_slots = getattr(self, "_send_slots", None)
                            if send_slots is not None:
                                send_slots.acquire()
                            try:
                                return self._send_one(
                                    log_id=item["log_id"],
                                    sender_snapshot=item["sender"],
                                    subject=subject,
                                    body_html=body_template,
                                    recipient_email=item["recipient_email"],
                                    recipient_name=item["recipient_name"],
                                    attachments=attachments,
                                    max_retries=0,
                                    retry_backoff_base=retry_backoff_base,
                                    proxy_url=item.get("proxy_url", ""),
                                )
                            finally:
                                if send_slots is not None:
                                    send_slots.release()
                        finally:
                            sender_slots.release()

                results: list[tuple[int, bool, str, int]] = []
                max_workers = min(worker_limit, len(assignments))
                if max_workers == 1:
                    results.append(execute_assignment(assignments[0]))
                else:
                    with ThreadPoolExecutor(max_workers=max_workers) as pool:
                        futures = [pool.submit(execute_assignment, item) for item in assignments]
                        for future in as_completed(futures):
                            results.append(future.result())

                batch_success = 0
                batch_failed = 0
                rate_limit_hits = 0
                progress_events = []
                task = session.get(Task, task_id)
                if not task:
                    return

                for log_id, success, error, sender_id in results:
                    log = session.get(SendLog, log_id)
                    sender = session.get(Sender, sender_id)
                    if not log:
                        continue
                    log.claimed_at = None
                    log.last_error_code = (error or "")[:80]
                    if success:
                        log.status = "success"
                        log.error_message = ""
                        log.sent_at = _utcnow()
                        log.next_attempt_at = None
                        batch_success += 1
                        if sender:
                            lb.report_success(
                                sender,
                                session,
                                count_attempt=False,
                                commit=False,
                            )
                        result_label = "success"
                    else:
                        log.error_message = (error or "")[:500]
                        if sender:
                            lb.report_failure(
                                sender,
                                session,
                                error,
                                count_attempt=False,
                                commit=False,
                            )

                        is_rl = is_rate_limit_error(error)
                        if not is_rl and error:
                            el = error.lower()
                            is_rl = any(p.lower() in el for p in rate_limit_patterns)

                        auth_error = is_auth_error(error)
                        if auth_error and sender:
                            sender.status = "banned"
                            session.add(sender)
                            logger.error("Sender %s banned due to auth failure", sender.email)

                        if is_rl:
                            rate_limit_hits += 1
                            log.status = "pending"
                            log.sent_at = None
                            log.next_attempt_at = _utcnow() + timedelta(seconds=rate_limit_cooldown)
                            if sender:
                                pause_for = max(rate_limit_cooldown, 30)
                                sender.status = "paused"
                                sender.paused_until = _utcnow() + timedelta(seconds=pause_for)
                                session.add(sender)
                            result_label = "retrying"
                        elif auth_error:
                            log.status = "failed"
                            log.sent_at = _utcnow()
                            log.next_attempt_at = None
                            batch_failed += 1
                            result_label = "failed"
                        elif log.attempt_count <= max_retries:
                            backoff = min(
                                3600.0,
                                retry_backoff_base ** max(0, log.attempt_count - 1),
                            )
                            log.status = "pending"
                            log.sent_at = None
                            log.next_attempt_at = _utcnow() + timedelta(seconds=backoff)
                            result_label = "retrying"
                        else:
                            log.status = "failed"
                            log.sent_at = _utcnow()
                            log.next_attempt_at = None
                            batch_failed += 1
                            result_label = "failed"

                    session.add(log)
                    progress_events.append((log.recipient_email, result_label, error))

                if rate_limit_hits:
                    consecutive_rate_limits += rate_limit_hits
                else:
                    consecutive_rate_limits = 0
                task.success_count = int(task.success_count or 0) + batch_success
                task.fail_count = int(task.fail_count or 0) + batch_failed
                task.last_heartbeat_at = _utcnow()
                session.add(task)
                session.commit()

                for recipient_email, result_label, error in progress_events:
                    publish_task_event(
                        task_id,
                        {
                            "type": "progress",
                            "status": task.status,
                            "success_count": task.success_count,
                            "fail_count": task.fail_count,
                            "recipient_count": task.recipient_count,
                            "last_recipient": recipient_email,
                            "last_result": result_label,
                            "error": (error or "")[:200],
                        },
                    )

                # Risk control: consecutive rate limits -> pause whole task
                if consecutive_rate_limits >= max_consecutive_rate_limits and risk_auto_pause_task:
                    wait_s = max(risk_pause_seconds, rate_limit_cooldown)
                    logger.warning(
                        "Task %s: risk detected (%s consecutive rate limits), pause %ss",
                        task_id,
                        consecutive_rate_limits,
                        wait_s,
                    )
                    task = session.get(Task, task_id)
                    if task:
                        self._pause_task(
                            task,
                            session,
                            reason="rate_limit",
                            delay_seconds=wait_s,
                            auto_resume=auto_resume_after_cooldown,
                        )
                    self._persist_circuit_breakers(lb, session)
                    publish_task_event(
                        task_id,
                        {
                            "type": "status",
                            "status": "paused",
                            "message": f"检测到风控，暂停 {wait_s}s 后自动恢复",
                            "success_count": task.success_count if task else 0,
                            "fail_count": task.fail_count if task else 0,
                            "recipient_count": task.recipient_count if task else 0,
                        },
                    )
                    return

            # Finalize counts
            session.refresh(task)
            if task.status == "running":
                counts = session.exec(
                    select(SendLog.status, func.count(SendLog.id))
                    .where(SendLog.task_id == task_id)
                    .group_by(SendLog.status)
                ).all()
                status_counts = {status: int(count) for status, count in counts}
                task.success_count = status_counts.get("success", 0)
                task.fail_count = status_counts.get("failed", 0)
                has_unfinished = bool(
                    status_counts.get("pending", 0) or status_counts.get("processing", 0)
                )
                if not has_unfinished:
                    task.status = "completed"
                    task.completed_at = _utcnow()
                session.add(task)
                session.commit()
                self._persist_circuit_breakers(lb, session)
                publish_task_event(
                    task_id,
                    {
                        "type": "status",
                        "status": task.status,
                        "success_count": task.success_count,
                        "fail_count": task.fail_count,
                        "recipient_count": task.recipient_count,
                        "message": "任务完成" if task.status == "completed" else f"任务状态: {task.status}",
                    },
                )
                logger.info(
                    "Task %s finished: success=%s fail=%s status=%s",
                    task_id,
                    task.success_count,
                    task.fail_count,
                    task.status,
                )

    def _restore_circuit_breakers(self, lb: LoadBalancer, senders: list[Sender], session: Session) -> None:
        for sender in senders:
            if sender.cb_state and sender.cb_state != "closed":
                with lb._lock:
                    cb = CircuitBreakerState()
                    cb.state = sender.cb_state
                    cb.failure_count = sender.cb_failure_count or 0
                    cb.success_count = sender.cb_success_count or 0
                    if sender.cb_next_attempt_time:
                        cb.next_attempt_time = to_unix_utc(sender.cb_next_attempt_time)
                    if sender.cb_last_failure_time:
                        cb.last_failure_time = to_unix_utc(sender.cb_last_failure_time)
                    lb._circuit_breakers[sender.id] = cb

    def _persist_circuit_breakers(self, lb: LoadBalancer, session: Session) -> None:
        with lb._lock:
            for sender_id, cb in lb._circuit_breakers.items():
                sender = session.get(Sender, sender_id)
                if not sender:
                    continue
                sender.cb_state = cb.state
                sender.cb_failure_count = cb.failure_count
                sender.cb_success_count = cb.success_count
                sender.cb_next_attempt_time = (
                    from_unix_utc(cb.next_attempt_time) if cb.next_attempt_time else None
                )
                sender.cb_last_failure_time = (
                    from_unix_utc(cb.last_failure_time) if cb.last_failure_time else None
                )
                session.add(sender)
            session.commit()

    def _check_scheduled_due(self, session: Session) -> None:
        """In-process scheduled task launcher so Celery is optional."""
        now = _utcnow()
        due = session.exec(
            select(Task).where(
                Task.status == "pending",
                Task.schedule_type == "scheduled",
                Task.schedule_time != None,  # noqa: E711
                Task.schedule_time <= now,
            ).order_by(Task.schedule_time).limit(500)
        ).all()
        for task in due:
            try:
                sender_ids = json.loads(task.sender_ids or "[]")
            except Exception:
                sender_ids = []
            senders = self._refresh_senders(session, sender_ids)
            if not senders:
                logger.warning("Scheduled task %s due but no senders available", task.id)
                continue
            if hasattr(task, "next_run_at"):
                task.next_run_at = None
                session.add(task)
                session.commit()
            self.submit(task.id)
            logger.info("Auto-started scheduled task %s", task.id)

        # smart schedule: start pending smart tasks immediately (best-effort)
        smart_tasks = session.exec(
            select(Task)
            .where(Task.status == "pending", Task.schedule_type == "smart")
            .order_by(Task.created_at)
            .limit(500)
        ).all()
        for task in smart_tasks:
            with self._lock:
                if task.id in self._active_tasks or task.id in self._queue:
                    continue
            self.submit(task.id)

    def _recover_once(self) -> None:
        with Session(engine) as session:
            self._check_scheduled_due(session)

            paused = session.exec(
                select(Task).where(Task.status == "paused").order_by(Task.id).limit(500)
            ).all()
            for task in paused:
                smart = _parse_smart_config(task.smart_config)
                if task.pause_reason == "manual" or not _coerce_bool(
                    smart.get("auto_resume_after_cooldown", True), True
                ):
                    continue
                pending = session.exec(
                    select(SendLog)
                    .where(SendLog.task_id == task.id)
                    .where(SendLog.status == "pending")
                    .limit(1)
                ).first()
                if not pending:
                    task.status = "completed"
                    task.completed_at = _utcnow()
                    session.add(task)
                    session.commit()
                    continue

                now_wall = time.time()
                now_utc = _utcnow()
                with self._lock:
                    if task.id in self._active_tasks or task.id in self._queue:
                        continue
                    resume_at = self._resume_at.get(task.id)
                    if resume_at and resume_at > now_wall:
                        continue
                    if resume_at:
                        self._resume_at.pop(task.id, None)

                if task.next_run_at and task.next_run_at > now_utc:
                    delay = max(1.0, (task.next_run_at - now_utc).total_seconds())
                    self.submit(task.id, delay_seconds=delay)
                else:
                    self.submit(task.id)

            running = session.exec(
                select(Task).where(Task.status == "running").order_by(Task.id).limit(500)
            ).all()
            for task in running:
                pending = session.exec(
                    select(SendLog)
                    .where(SendLog.task_id == task.id)
                    .where(SendLog.status == "pending")
                    .limit(1)
                ).first()
                if pending:
                    with self._lock:
                        if task.id not in self._active_tasks and task.id not in self._queue:
                            self._queue.append(task.id)
                            logger.info("Recovered stuck task %s", task.id)

    def _recovery_loop(self) -> None:
        while self._running and not self._stop_event.wait(10):
            try:
                self._recover_once()
            except Exception:
                logger.exception("Recovery scan failed")

    def shutdown(self, timeout: Optional[float] = None) -> None:
        self._running = False
        self._stop_event.set()
        wait_timeout = timeout
        if wait_timeout is None:
            wait_timeout = max(5.0, float(getattr(settings, "SMTP_TIMEOUT", 30)) + 5.0)
        deadline = time.monotonic() + max(0.0, wait_timeout)
        current = threading.current_thread()
        threads = list(self._workers)
        recovery_thread = getattr(self, "_recovery_thread", None)
        if recovery_thread is not None:
            threads.append(recovery_thread)
        for thread in threads:
            if thread is current or not thread.is_alive():
                continue
            thread.join(timeout=max(0.0, deadline - time.monotonic()))
        alive = [thread for thread in threads if thread is not current and thread.is_alive()]
        if not alive:
            try:
                email_sender.close()
            except Exception:
                logger.exception("Failed to close SMTP connection pool")
        else:
            logger.warning("SendEngine shutdown timed out with %s worker(s) active", len(alive))


def get_send_engine() -> SendEngine:
    global _engine_instance
    with _engine_lock:
        if _engine_instance is None:
            workers = 4
            try:
                from ..config import settings

                workers = int(getattr(settings, "SEND_ENGINE_WORKERS", 4) or 4)
            except (TypeError, ValueError):
                logger.warning("Invalid SEND_ENGINE_WORKERS value; using 4")
            _engine_instance = SendEngine(max_workers=max(1, min(32, workers)))
        return _engine_instance
