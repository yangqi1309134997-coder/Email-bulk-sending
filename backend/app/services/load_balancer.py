"""Load balancer with round-robin / weighted / smart strategies and circuit breaker."""

from __future__ import annotations

import random
import threading
import time
from dataclasses import dataclass
from datetime import timedelta
from typing import List, Optional

from ..models.sender import Sender
from ..utils.time import utcnow


@dataclass
class CircuitBreakerState:
    failure_count: int = 0
    success_count: int = 0
    last_failure_time: float = 0
    state: str = "closed"  # closed, open, half-open
    next_attempt_time: float = 0


class LoadBalancer:
    """负载均衡器：支持轮询、权重、智能三种策略 + 熔断器模式。"""

    def __init__(
        self,
        strategy: str = "round_robin",
        failure_threshold: int = 5,
        success_threshold: int = 2,
        timeout: int = 60,
        consecutive_fail_pause: int = 3,
        pause_minutes: int = 5,
    ):
        self.strategy = strategy or "round_robin"
        self._round_robin_index = 0
        self._lock = threading.RLock()
        self._circuit_breakers: dict[int, CircuitBreakerState] = {}
        self.failure_threshold = failure_threshold
        self.success_threshold = success_threshold
        self.timeout = timeout
        self.consecutive_fail_pause = consecutive_fail_pause
        self.pause_minutes = pause_minutes

    def pick_sender(self, senders: List[Sender]) -> Optional[Sender]:
        available = [s for s in senders if s.is_available() and self._is_circuit_closed(s.id)]
        if not available:
            return None

        if self.strategy == "weighted":
            return self._weighted(available)
        if self.strategy == "smart":
            return self._smart(available)
        return self._round_robin(available)

    def _is_circuit_closed(self, sender_id: int) -> bool:
        with self._lock:
            cb = self._circuit_breakers.get(sender_id)
            if not cb:
                return True
            now = time.time()
            if cb.state == "open":
                if now >= cb.next_attempt_time:
                    cb.state = "half-open"
                    cb.success_count = 0
                    return True
                return False
            return True

    def _round_robin(self, senders: List[Sender]) -> Sender:
        with self._lock:
            if not senders:
                raise ValueError("empty senders")
            sender = senders[self._round_robin_index % len(senders)]
            self._round_robin_index += 1
            return sender

    def _weighted(self, senders: List[Sender]) -> Sender:
        weights = [max(1, int(s.weight or 1)) for s in senders]
        # Selection randomness is operational, not security-sensitive.
        return random.choices(senders, weights=weights, k=1)[0]  # nosec B311

    def _smart(self, senders: List[Sender]) -> Sender:
        scores = []
        for s in senders:
            if s.daily_quota and s.daily_quota > 0:
                quota_remaining = max(0.0, (s.daily_quota - s.daily_sent) / s.daily_quota)
            else:
                quota_remaining = 0.5
            success_rate = s.success_rate if s.success_rate is not None else 1.0
            success_rate = max(0.0, min(1.0, float(success_rate)))
            cb = self._circuit_breakers.get(s.id)
            cb_penalty = 0.5 if cb and cb.state == "half-open" else 1.0
            # Prefer higher remaining quota and higher success rate; slight weight bias
            weight_factor = max(0.1, (s.weight or 50) / 100.0)
            score = (quota_remaining * 0.35 + success_rate * 0.5 + weight_factor * 0.15) * cb_penalty
            scores.append(max(0.01, score))
        # Selection randomness is operational, not security-sensitive.
        return random.choices(senders, weights=scores, k=1)[0]  # nosec B311

    def report_success(
        self,
        sender: Sender,
        session,
        *,
        count_attempt: bool = True,
        commit: bool = True,
    ) -> None:
        if count_attempt:
            sender.daily_sent = int(sender.daily_sent or 0) + 1
        sender.consecutive_failures = 0
        total = max(1, int(sender.daily_sent))
        prev_rate = float(sender.success_rate if sender.success_rate is not None else 1.0)
        # EWMA-ish update based on cumulative attempts approximation
        sender.success_rate = (prev_rate * (total - 1) + 1.0) / total
        if sender.status == "paused" and sender.paused_until and sender.paused_until <= utcnow():
            sender.status = "active"
            sender.paused_until = None
        session.add(sender)
        if commit:
            session.commit()

        with self._lock:
            cb = self._circuit_breakers.get(sender.id)
            if not cb:
                cb = CircuitBreakerState()
                self._circuit_breakers[sender.id] = cb
            cb.failure_count = 0
            if cb.state == "half-open":
                cb.success_count += 1
                if cb.success_count >= self.success_threshold:
                    cb.state = "closed"
                    cb.success_count = 0

    def report_failure(
        self,
        sender: Sender,
        session,
        error: str = "",
        *,
        count_attempt: bool = True,
        commit: bool = True,
    ) -> None:
        sender.consecutive_failures = int(sender.consecutive_failures or 0) + 1
        if count_attempt:
            sender.daily_sent = int(sender.daily_sent or 0) + 1
        total = max(1, int(sender.daily_sent))
        prev_rate = float(sender.success_rate if sender.success_rate is not None else 1.0)
        sender.success_rate = (prev_rate * (total - 1)) / total

        if sender.consecutive_failures >= self.consecutive_fail_pause:
            sender.status = "paused"
            sender.paused_until = utcnow() + timedelta(minutes=self.pause_minutes)

        session.add(sender)
        if commit:
            session.commit()

        with self._lock:
            cb = self._circuit_breakers.get(sender.id)
            if not cb:
                cb = CircuitBreakerState()
                self._circuit_breakers[sender.id] = cb
            cb.failure_count += 1
            cb.last_failure_time = time.time()
            if cb.state == "half-open":
                cb.state = "open"
                cb.next_attempt_time = time.time() + self.timeout
            elif cb.failure_count >= self.failure_threshold:
                cb.state = "open"
                cb.next_attempt_time = time.time() + self.timeout

    def get_circuit_state(self, sender_id: int) -> dict:
        with self._lock:
            cb = self._circuit_breakers.get(sender_id)
            if not cb:
                return {"state": "closed", "failure_count": 0, "success_count": 0}
            return {
                "state": cb.state,
                "failure_count": cb.failure_count,
                "success_count": cb.success_count,
                "next_attempt": cb.next_attempt_time if cb.state == "open" else None,
            }

    def reset_circuit(self, sender_id: int) -> None:
        with self._lock:
            self._circuit_breakers.pop(sender_id, None)
