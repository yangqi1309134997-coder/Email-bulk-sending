"""Tests for load balancer strategies and circuit breaker."""

from types import SimpleNamespace
from unittest.mock import MagicMock

from app.services.load_balancer import LoadBalancer


def make_sender(id=1, weight=50, daily_quota=100, daily_sent=0, success_rate=1.0, status="active", enabled=True):
    return SimpleNamespace(
        id=id,
        weight=weight,
        daily_quota=daily_quota,
        daily_sent=daily_sent,
        success_rate=success_rate,
        status=status,
        enabled=enabled,
        consecutive_failures=0,
        paused_until=None,
        is_available=lambda self=None: True,
    )


def test_round_robin_cycles():
    lb = LoadBalancer(strategy="round_robin")
    s1 = make_sender(1)
    s2 = make_sender(2)
    # override is_available bound correctly
    for s in (s1, s2):
        s.is_available = lambda: True
    picks = [lb.pick_sender([s1, s2]).id for _ in range(4)]
    assert picks == [1, 2, 1, 2]


def test_weighted_prefers_higher_weight():
    lb = LoadBalancer(strategy="weighted")
    s1 = make_sender(1, weight=1)
    s2 = make_sender(2, weight=1000)
    s1.is_available = lambda: True
    s2.is_available = lambda: True
    picks = [lb.pick_sender([s1, s2]).id for _ in range(50)]
    assert picks.count(2) > picks.count(1)


def test_circuit_opens_after_threshold_failures():
    lb = LoadBalancer(strategy="round_robin", failure_threshold=3, timeout=60)
    sender = make_sender(1)
    sender.is_available = lambda: True
    session = MagicMock()

    for _ in range(3):
        lb.report_failure(sender, session, "error")

    state = lb.get_circuit_state(1)
    assert state["state"] == "open"
    # open circuit excludes sender
    assert lb.pick_sender([sender]) is None


def test_report_success_closes_half_open():
    lb = LoadBalancer(strategy="round_robin", success_threshold=2)
    sender = make_sender(1)
    sender.is_available = lambda: True
    session = MagicMock()

    # force half-open
    with lb._lock:
        from app.services.load_balancer import CircuitBreakerState
        cb = CircuitBreakerState(state="half-open", success_count=0)
        lb._circuit_breakers[1] = cb

    lb.report_success(sender, session)
    lb.report_success(sender, session)
    assert lb.get_circuit_state(1)["state"] == "closed"
