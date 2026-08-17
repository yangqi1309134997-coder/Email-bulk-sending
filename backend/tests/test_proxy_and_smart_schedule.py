"""Tests for proxy parsing, smart window selection, and event publishing."""

from datetime import datetime

from app.services.email_sender import _normalize_smtp_host, _parse_proxy
from app.services.send_engine import (
    _is_good_send_hour,
    _pick_proxy,
    _recipient_timezone_offset,
    publish_task_event,
)
from app.websocket import events as ws_events


def test_parse_proxy_http_with_auth():
    info = _parse_proxy("http://user:pass@1.2.3.4:8080")
    assert info is not None
    assert info["host"] == "1.2.3.4"
    assert info["port"] == 8080
    assert info["username"] == "user"
    assert info["password"] == "pass"


def test_parse_proxy_decodes_auth_and_rejects_url_paths():
    info = _parse_proxy("http://user%40name:p%40ss@proxy.example.com:8443")
    assert info is not None
    assert info["username"] == "user@name"
    assert info["password"] == "p@ss"
    assert _parse_proxy("http://proxy.example.com:8080/unexpected") is None


def test_smtp_host_validation_blocks_proxy_header_injection():
    assert _normalize_smtp_host("SMTP.Example.COM.") == "smtp.example.com"
    assert _normalize_smtp_host("[2001:db8::1]") == "2001:db8::1"
    for invalid in ("smtp.example.com\r\nX-Test: yes", "https://smtp.example.com", "host:25"):
        try:
            _normalize_smtp_host(invalid)
        except ValueError:
            pass
        else:
            raise AssertionError(f"invalid SMTP host accepted: {invalid!r}")


def test_parse_proxy_plain_host_port():
    info = _parse_proxy("10.0.0.8:3128")
    assert info is not None
    assert info["host"] == "10.0.0.8"
    assert info["port"] == 3128


def test_parse_proxy_invalid():
    assert _parse_proxy("") is None
    assert _parse_proxy("not-a-proxy") is None


def test_pick_proxy_round_robin():
    proxies = ["http://a:1", "http://b:2"]
    assert _pick_proxy(proxies, 0) == "http://a:1"
    assert _pick_proxy(proxies, 1) == "http://b:2"
    assert _pick_proxy(proxies, 2) == "http://a:1"
    assert _pick_proxy([], 0) == ""


def test_recipient_timezone_offset_known_domains():
    assert _recipient_timezone_offset("a@qq.com") == 8
    assert _recipient_timezone_offset("a@gmail.com") == -5
    assert _recipient_timezone_offset("a@company.de") == 1


def test_is_good_send_hour_china_morning():
    # UTC 01:00 -> CST 09:00
    now = datetime(2026, 7, 17, 1, 0, 0)
    assert _is_good_send_hour("user@163.com", now) is True
    # UTC 04:00 -> CST 12:00 not in preferred windows
    now2 = datetime(2026, 7, 17, 4, 0, 0)
    assert _is_good_send_hour("user@163.com", now2) is False


def test_publish_task_event_enqueues():
    while True:
        try:
            ws_events._event_queue.get_nowait()
        except Exception:
            break
    publish_task_event(123, {"type": "progress", "status": "running"})
    task_id, message = ws_events._event_queue.get_nowait()
    assert task_id == 123
    assert message["type"] == "progress"
    assert message["task_id"] == 123
