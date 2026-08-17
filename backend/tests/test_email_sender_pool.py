"""Unit tests for SMTP connection pool and email sender helpers."""

import io
import smtplib
import threading
import time
from types import SimpleNamespace

import pytest

from app.services.email_sender import (
    EmailSender,
    SMTPConnectionPool,
    is_auth_error,
    is_rate_limit_error,
)


class FakeSMTP:
    def __init__(self, host, port, timeout=None, context=None):
        self.host = host
        self.port = port
        self.timeout = timeout
        self.context = context
        self.closed = False
        self.starttls_calls = 0
        self.login_calls = []
        self.ehlo_calls = 0

    def ehlo(self):
        self.ehlo_calls += 1
        return 250, b"ok"

    def starttls(self, context=None):
        self.starttls_calls += 1
        self.context = context

    def login(self, username, password):
        self.login_calls.append((username, password))

    def noop(self):
        if self.closed:
            raise smtplib.SMTPServerDisconnected("closed")
        return 250, b"ok"

    def quit(self):
        self.closed = True

    def close(self):
        self.closed = True


def make_sender(**overrides):
    values = {
        "id": 1,
        "email": "sender@example.com",
        "password": "encrypted",
        "smtp_server": "smtp.example.com",
        "smtp_port": 587,
        "use_tls": True,
        "sender_type": "自定义SMTP",
    }
    values.update(overrides)
    return SimpleNamespace(**values)


def test_pool_uses_starttls_for_587(monkeypatch):
    created = []

    def create_smtp(*args, **kwargs):
        server = FakeSMTP(*args, **kwargs)
        created.append(server)
        return server

    monkeypatch.setattr(smtplib, "SMTP", create_smtp)
    monkeypatch.setattr(smtplib, "SMTP_SSL", lambda *a, **k: pytest.fail("SSL not for 587"))
    monkeypatch.setattr("app.services.email_sender.decrypt_password", lambda value: "secret")

    pool = SMTPConnectionPool(max_connections_per_sender=1)
    sender = make_sender(smtp_port=587, use_tls=True)

    with pool.get_connection(sender) as connection:
        assert connection is created[0]
        assert created[0].starttls_calls == 1
        assert created[0].login_calls == [("sender@example.com", "secret")]
    assert pool._creating == {}


def test_pool_uses_implicit_ssl_for_465(monkeypatch):
    created = []

    def create_ssl(*args, **kwargs):
        server = FakeSMTP(*args, **kwargs)
        created.append(server)
        return server

    monkeypatch.setattr(smtplib, "SMTP_SSL", create_ssl)
    monkeypatch.setattr(smtplib, "SMTP", lambda *a, **k: pytest.fail("plain SMTP must not be used for 465"))
    monkeypatch.setattr("app.services.email_sender.decrypt_password", lambda value: "secret")

    pool = SMTPConnectionPool(max_connections_per_sender=1)
    sender = make_sender(smtp_port=465, use_tls=True)

    with pool.get_connection(sender) as connection:
        assert connection is created[0]
        assert created[0].login_calls == [("sender@example.com", "secret")]


def test_pool_never_exceeds_the_configured_connection_limit(monkeypatch):
    created = []
    created_lock = threading.Lock()

    def create_smtp(*args, **kwargs):
        server = FakeSMTP(*args, **kwargs)
        with created_lock:
            created.append(server)
        return server

    monkeypatch.setattr(smtplib, "SMTP", create_smtp)
    monkeypatch.setattr("app.services.email_sender.decrypt_password", lambda value: "secret")

    pool = SMTPConnectionPool(max_connections_per_sender=2, acquire_timeout=0.5)
    sender = make_sender()
    barrier = threading.Barrier(8)
    failures = []

    def use_connection():
        try:
            barrier.wait(timeout=2)
            with pool.get_connection(sender):
                time.sleep(0.05)
        except Exception as exc:  # pragma: no cover
            failures.append(exc)

    threads = [threading.Thread(target=use_connection) for _ in range(8)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=5)

    assert failures == []
    assert len(created) == 2


def test_pool_does_not_reuse_authenticated_connection_for_different_sender(monkeypatch):
    created = []

    def create_smtp(*args, **kwargs):
        server = FakeSMTP(*args, **kwargs)
        created.append(server)
        return server

    monkeypatch.setattr(smtplib, "SMTP", create_smtp)
    monkeypatch.setattr("app.services.email_sender.decrypt_password", lambda value: value)
    pool = SMTPConnectionPool(max_connections_per_sender=1)

    first = make_sender(id=1, password="credential-a")
    second = make_sender(id=2, password="credential-b")
    with pool.get_connection(first):
        pass
    with pool.get_connection(second):
        pass

    assert len(created) == 2
    assert created[0].login_calls == [("sender@example.com", "credential-a")]
    assert created[1].login_calls == [("sender@example.com", "credential-b")]


def test_slow_connect_does_not_block_unrelated_sender(monkeypatch):
    active = 0
    max_active = 0
    active_lock = threading.Lock()

    def create_smtp(*args, **kwargs):
        nonlocal active, max_active
        with active_lock:
            active += 1
            max_active = max(max_active, active)
        time.sleep(0.05)
        with active_lock:
            active -= 1
        return FakeSMTP(*args, **kwargs)

    monkeypatch.setattr(smtplib, "SMTP", create_smtp)
    monkeypatch.setattr("app.services.email_sender.decrypt_password", lambda value: value)
    pool = SMTPConnectionPool(max_connections_per_sender=1)
    barrier = threading.Barrier(2)

    def acquire(sender):
        barrier.wait(timeout=1)
        with pool.get_connection(sender):
            pass

    senders = [
        make_sender(id=1, email="one@example.com", smtp_server="smtp.one.example"),
        make_sender(id=2, email="two@example.com", smtp_server="smtp.two.example"),
    ]
    threads = [threading.Thread(target=acquire, args=(sender,)) for sender in senders]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=2)

    assert max_active == 2


def test_broken_connection_is_removed_from_pool(monkeypatch):
    created = []

    def create_smtp(*args, **kwargs):
        server = FakeSMTP(*args, **kwargs)
        created.append(server)
        return server

    monkeypatch.setattr(smtplib, "SMTP", create_smtp)
    monkeypatch.setattr("app.services.email_sender.decrypt_password", lambda value: "secret")

    pool = SMTPConnectionPool(max_connections_per_sender=1)
    sender = make_sender()

    with pytest.raises(RuntimeError):
        with pool.get_connection(sender):
            raise RuntimeError("send failed")

    with pool.get_connection(sender):
        pass

    assert len(created) == 2
    assert created[0].closed is True


def test_invalid_proxy_never_falls_back_to_direct_smtp(monkeypatch):
    monkeypatch.setattr("app.services.email_sender.decrypt_password", lambda value: "secret")
    pool = SMTPConnectionPool(max_connections_per_sender=1)
    sender = make_sender()

    with pytest.raises(ValueError, match="proxy"):
        pool._create_connection(sender, proxy_url="http://proxy.example.com:8080/path")


def test_proxy_socket_consumes_smtp_greeting_before_ehlo():
    class GreetingSocket:
        def __init__(self, response):
            self.response = response
            self.closed = False

        def makefile(self, mode):
            return io.BytesIO(self.response)

        def close(self):
            self.closed = True

    pool = SMTPConnectionPool()
    server = smtplib.SMTP()
    sock = GreetingSocket(b"220 smtp.example.com ready\r\n")

    attached = pool._attach_connected_socket(server, sock, "smtp.example.com")

    assert attached is server
    assert server._host == "smtp.example.com"


def test_recipient_variables_are_html_escaped_in_message_body():
    message = EmailSender().build_message(
        "sender@example.com",
        "recipient@example.com",
        '<img src=x onerror="alert(1)">',
        "Hello {name}",
        "<p>{name} / {email}</p>",
    )

    body = message.get_payload()[0].get_payload(decode=True).decode("utf-8")
    assert "&lt;img" in body
    assert "<img src=x" not in body


def test_rate_limit_and_auth_detection():
    assert is_rate_limit_error("421 Too many attempts")
    assert is_rate_limit_error("rate limit exceeded")
    assert is_rate_limit_error("spam blocked")
    assert not is_rate_limit_error("mailbox full random")
    assert not is_rate_limit_error("552 message size exceeds fixed maximum")
    assert not is_rate_limit_error("554 permanent transaction failure")
    assert is_auth_error("535 Authentication failed")
    assert is_auth_error("Auth failed: invalid credentials")
