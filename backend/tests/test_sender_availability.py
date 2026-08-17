"""Tests for Sender.is_available and pause expiry logic."""

from datetime import timedelta

from app.models.sender import Sender
from app.utils.time import utcnow


def test_available_active():
    s = Sender(
        user_id=1,
        email="a@b.com",
        password="x",
        smtp_server="smtp.example.com",
        enabled=True,
        status="active",
        daily_quota=100,
        daily_sent=10,
    )
    assert s.is_available() is True


def test_unavailable_when_quota_exhausted():
    s = Sender(
        user_id=1,
        email="a@b.com",
        password="x",
        smtp_server="smtp.example.com",
        enabled=True,
        status="active",
        daily_quota=10,
        daily_sent=10,
    )
    assert s.is_available() is False


def test_paused_until_future_unavailable():
    s = Sender(
        user_id=1,
        email="a@b.com",
        password="x",
        smtp_server="smtp.example.com",
        enabled=True,
        status="paused",
        paused_until=utcnow() + timedelta(minutes=10),
        daily_quota=100,
        daily_sent=0,
    )
    assert s.is_available() is False


def test_paused_expired_available():
    s = Sender(
        user_id=1,
        email="a@b.com",
        password="x",
        smtp_server="smtp.example.com",
        enabled=True,
        status="paused",
        paused_until=utcnow() - timedelta(minutes=1),
        daily_quota=100,
        daily_sent=0,
    )
    assert s.is_available() is True


def test_banned_unavailable():
    s = Sender(
        user_id=1,
        email="a@b.com",
        password="x",
        smtp_server="smtp.example.com",
        enabled=True,
        status="banned",
        daily_quota=100,
        daily_sent=0,
    )
    assert s.is_available() is False


def test_paused_without_until_unavailable():
    s = Sender(
        user_id=1,
        email="a@b.com",
        password="x",
        smtp_server="smtp.example.com",
        enabled=True,
        status="paused",
        paused_until=None,
        daily_quota=100,
        daily_sent=0,
    )
    assert s.is_available() is False
