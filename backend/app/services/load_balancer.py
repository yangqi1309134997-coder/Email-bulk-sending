import random
from typing import List, Optional
from ..models.sender import Sender


class LoadBalancer:
    """负载均衡器：支持轮询、权重、智能三种策略"""

    def __init__(self, strategy: str = "round_robin"):
        self.strategy = strategy
        self._round_robin_index = 0

    def pick_sender(self, senders: List[Sender]) -> Optional[Sender]:
        available = [s for s in senders if s.is_available()]
        if not available:
            return None

        if self.strategy == "round_robin":
            return self._round_robin(available)
        elif self.strategy == "weighted":
            return self._weighted(available)
        elif self.strategy == "smart":
            return self._smart(available)
        else:
            return self._round_robin(available)

    def _round_robin(self, senders: List[Sender]) -> Sender:
        sender = senders[self._round_robin_index % len(senders)]
        self._round_robin_index += 1
        return sender

    def _weighted(self, senders: List[Sender]) -> Sender:
        weights = [s.weight for s in senders]
        return random.choices(senders, weights=weights, k=1)[0]

    def _smart(self, senders: List[Sender]) -> Sender:
        scores = []
        for s in senders:
            quota_remaining = (s.daily_quota - s.daily_sent) / s.daily_quota if s.daily_quota > 0 else 0
            score = quota_remaining * 0.4 + s.success_rate * 0.6
            scores.append(score)
        return random.choices(senders, weights=scores, k=1)[0]

    def report_success(self, sender: Sender, session):
        sender.daily_sent += 1
        sender.consecutive_failures = 0
        # Update success rate
        total = sender.daily_sent
        if total > 0:
            sender.success_rate = (sender.success_rate * (total - 1) + 1) / total
        session.add(sender)
        session.commit()

    def report_failure(self, sender: Sender, session):
        sender.consecutive_failures += 1
        sender.daily_sent += 1
        # Update success rate
        total = sender.daily_sent
        if total > 0:
            sender.success_rate = (sender.success_rate * (total - 1)) / total

        # Smart degradation: pause after 3 consecutive failures
        if sender.consecutive_failures >= 3:
            from datetime import datetime, timedelta
            sender.status = "paused"
            sender.paused_until = datetime.utcnow() + timedelta(minutes=5)

        session.add(sender)
        session.commit()