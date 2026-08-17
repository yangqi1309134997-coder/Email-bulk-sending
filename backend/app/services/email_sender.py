"""SMTP email sender with connection pooling, SSL/TLS support, and retries."""

from __future__ import annotations

import logging
import os
import random
import hashlib
import html
import ipaddress
import re
import smtplib
import ssl
import threading
import time
from collections import OrderedDict
from contextlib import contextmanager
from dataclasses import dataclass, field
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from typing import Generator, List, Optional, Tuple
from urllib.parse import unquote, urlparse

from ..config import settings
from ..models.sender import Sender
from ..utils.security import decrypt_password

logger = logging.getLogger(__name__)
_SUPPORTED_PROXY_SCHEMES = {"http", "socks4", "socks4a", "socks5", "socks5h"}


def _parse_proxy(proxy_url: str) -> Optional[dict]:
    """Parse http://user:pass@host:port or socks5://host:port into host/port/auth."""
    if not proxy_url:
        return None
    raw = proxy_url.strip()
    if not raw:
        return None
    if any(ord(char) < 32 for char in raw) or len(raw) > 2048:
        return None
    try:
        if "://" not in raw:
            raw = "http://" + raw
        parsed = urlparse(raw)
        scheme = (parsed.scheme or "http").lower()
        if scheme not in _SUPPORTED_PROXY_SCHEMES:
            return None
        if parsed.path not in ("", "/") or parsed.params or parsed.query or parsed.fragment:
            return None
        if not parsed.hostname or parsed.port is None or not (1 <= int(parsed.port) <= 65535):
            return None
        return {
            "scheme": scheme,
            "host": parsed.hostname,
            "port": int(parsed.port),
            "username": unquote(parsed.username or ""),
            "password": unquote(parsed.password or ""),
            "url": raw,
        }
    except Exception:
        return None


def _normalize_smtp_host(host: str) -> str:
    """Validate a host before it reaches SMTP or a proxy request line."""
    value = str(host or "").strip()
    if not value or len(value) > 253 or any(ord(char) < 33 for char in value):
        raise ValueError("Invalid SMTP server")
    if any(char in value for char in "/?#@"):
        raise ValueError("Invalid SMTP server")
    if value.startswith("[") and value.endswith("]"):
        value = value[1:-1]
    try:
        ipaddress.ip_address(value)
        return value
    except ValueError:
        pass
    if ":" in value:
        raise ValueError("Invalid SMTP server")
    normalized = value.rstrip(".").lower()
    labels = normalized.split(".")
    label_pattern = re.compile(r"^[a-z0-9_](?:[a-z0-9_-]{0,61}[a-z0-9_])?$", re.IGNORECASE)
    if any(not label or len(label) > 63 or not label_pattern.fullmatch(label) for label in labels):
        raise ValueError("Invalid SMTP server")
    return normalized




@dataclass
class PooledConnection:
    server: smtplib.SMTP
    sender_email: str
    in_use: bool = False
    created_at: float = field(default_factory=time.time)
    last_used: float = field(default_factory=time.time)
    use_count: int = 0
    is_temporary: bool = False


class SMTPConnectionPool:
    """Thread-safe SMTP connection pool with SSL/STARTTLS and idle cleanup."""

    def __init__(
        self,
        max_connections_per_sender: int = 5,
        max_idle_time: int = 300,
        acquire_timeout: float = 30.0,
        max_uses_per_connection: int = 100,
    ):
        self._pools: dict[str, list[PooledConnection]] = {}
        self._creating: dict[str, int] = {}
        self._lock = threading.RLock()
        self._condition = threading.Condition(self._lock)
        self.max_connections_per_sender = max_connections_per_sender
        self.max_idle_time = max_idle_time
        self.acquire_timeout = acquire_timeout
        self.max_uses_per_connection = max_uses_per_connection
        self._cleanup_interval = 60
        self._last_cleanup = time.time()

    def _get_pool_key(self, sender: Sender, proxy_url: str = "") -> str:
        username = getattr(sender, "smtp_username", "") or sender.email
        identity = getattr(sender, "id", None) or "unsaved"
        credential = hashlib.sha256(str(sender.password or "").encode("utf-8")).hexdigest()[:16]
        security = getattr(sender, "smtp_security", "") or "auto"
        smtp_host = _normalize_smtp_host(sender.smtp_server)
        return (
            f"{identity}:{smtp_host}:{sender.smtp_port}:{username}:"
            f"{security}:{bool(sender.use_tls)}:{credential}:{proxy_url or '-'}"
        )

    def _uses_implicit_ssl(self, sender: Sender) -> bool:
        security = (getattr(sender, "smtp_security", "") or "").lower()
        if security == "ssl":
            return True
        if security in ("starttls", "none"):
            return False
        return int(sender.smtp_port or 0) == 465

    def _uses_starttls(self, sender: Sender) -> bool:
        """Resolve the explicit security mode before falling back to legacy flags."""
        security = (getattr(sender, "smtp_security", "") or "").lower()
        if security in {"ssl", "none"}:
            return False
        if security == "starttls":
            return True
        return bool(sender.use_tls) and not self._uses_implicit_ssl(sender)

    @staticmethod
    def _attach_connected_socket(
        server: smtplib.SMTP,
        sock,
        smtp_host: str,
    ) -> smtplib.SMTP:
        """Initialize smtplib state for a socket opened through a proxy."""
        server.sock = sock
        server.file = sock.makefile("rb")
        server._host = smtp_host
        code, message = server.getreply()
        if code != 220:
            server.close()
            raise smtplib.SMTPConnectError(code, message)
        return server

    def _create_connection(self, sender: Sender, proxy_url: str = "") -> smtplib.SMTP:
        timeout = getattr(settings, "SMTP_TIMEOUT", 30)
        if int(sender.smtp_port or 0) <= 0:
            raise ValueError("SMTP port must be greater than zero")
        smtp_host = _normalize_smtp_host(sender.smtp_server)
        pwd = decrypt_password(sender.password)
        context = ssl.create_default_context()
        proxy = _parse_proxy(proxy_url)
        if str(proxy_url or "").strip() and proxy is None:
            raise ValueError("Invalid proxy URL")

        if proxy:
            # HTTP CONNECT tunnel for SMTP. Requires PySocks-compatible socket if socks*, else pure HTTP CONNECT.
            import socket
            scheme = proxy["scheme"]
            if scheme.startswith("socks"):
                try:
                    import socks  # type: ignore
                except Exception as exc:  # pragma: no cover
                    raise RuntimeError("SOCKS proxy requires PySocks package") from exc
                sock = socks.socksocket()
                proxy_type = socks.SOCKS5 if "5" in scheme else socks.SOCKS4
                sock.set_proxy(
                    proxy_type,
                    proxy["host"],
                    proxy["port"],
                    rdns=scheme in {"socks4a", "socks5h"},
                    username=proxy["username"] or None,
                    password=proxy["password"] or None,
                )
                sock.settimeout(timeout)
                sock.connect((smtp_host, int(sender.smtp_port)))
            else:
                # HTTP proxy CONNECT
                sock = socket.create_connection((proxy["host"], proxy["port"]), timeout=timeout)
                connect_host = smtp_host
                try:
                    if ipaddress.ip_address(smtp_host).version == 6:
                        connect_host = f"[{smtp_host}]"
                except ValueError:
                    pass
                connect_req = (
                    f"CONNECT {connect_host}:{int(sender.smtp_port)} HTTP/1.1\r\n"
                    f"Host: {connect_host}:{int(sender.smtp_port)}\r\n"
                )
                if proxy["username"]:
                    import base64
                    token = base64.b64encode(f"{proxy['username']}:{proxy['password']}".encode()).decode()
                    connect_req += f"Proxy-Authorization: Basic {token}\r\n"
                connect_req += "\r\n"
                sock.sendall(connect_req.encode())
                # Read proxy response headers
                resp = b""
                while b"\r\n\r\n" not in resp and len(resp) < 4096:
                    # Read only through the CONNECT header terminator. A
                    # larger recv can consume the tunneled SMTP 220 greeting
                    # and leave smtplib's protocol state permanently offset.
                    chunk = sock.recv(1)
                    if not chunk:
                        break
                    resp += chunk
                status_line = resp.split(b"\r\n", 1)[0].decode(errors="ignore")
                status_parts = status_line.split(" ", 2)
                if (
                    b"\r\n\r\n" not in resp
                    or len(status_parts) < 2
                    or status_parts[1] != "200"
                ):
                    sock.close()
                    raise RuntimeError(f"Proxy CONNECT failed: {status_line or 'empty response'}")

            if self._uses_implicit_ssl(sender):
                server = smtplib.SMTP_SSL()
                tls_sock = context.wrap_socket(sock, server_hostname=smtp_host)
                self._attach_connected_socket(server, tls_sock, smtp_host)
            else:
                server = smtplib.SMTP()
                self._attach_connected_socket(server, sock, smtp_host)
                if self._uses_starttls(sender):
                    server.ehlo()
                    server.starttls(context=context)
                    server.ehlo()
        else:
            if self._uses_implicit_ssl(sender):
                server = smtplib.SMTP_SSL(
                    smtp_host,
                    sender.smtp_port,
                    timeout=timeout,
                    context=context,
                )
            else:
                server = smtplib.SMTP(smtp_host, sender.smtp_port, timeout=timeout)
                if self._uses_starttls(sender):
                    server.ehlo()
                    server.starttls(context=context)
                    server.ehlo()

        login_user = getattr(sender, "smtp_username", None) or sender.email
        # Enterprise SMTP relays often use API key as username
        if sender.sender_type in ("SendGrid", "sendgrid"):
            login_user = "apikey"
        server.login(login_user, pwd)
        return server

    def _close_connection(self, pc: PooledConnection) -> None:
        try:
            pc.server.quit()
        except Exception:
            try:
                pc.server.close()
            except Exception:
                logger.debug("SMTP socket close failed", exc_info=True)

    def _finish_creation_unlocked(self, pool_key: str) -> None:
        remaining = max(0, self._creating.get(pool_key, 1) - 1)
        if remaining:
            self._creating[pool_key] = remaining
        else:
            self._creating.pop(pool_key, None)

    def _cleanup_idle_unlocked(self) -> list[PooledConnection]:
        now = time.time()
        self._last_cleanup = now
        discarded: list[PooledConnection] = []
        for pool_key, pool in list(self._pools.items()):
            to_remove = [
                pc
                for pc in pool
                if not pc.in_use
                and (
                    (now - pc.last_used) > self.max_idle_time
                    or pc.use_count >= self.max_uses_per_connection
                )
            ]
            for pc in to_remove:
                pool.remove(pc)
                discarded.append(pc)
            if not pool:
                del self._pools[pool_key]
        return discarded

    @contextmanager
    def get_connection(self, sender: Sender, proxy_url: str = "") -> Generator[smtplib.SMTP, None, None]:
        pool_key = self._get_pool_key(sender, proxy_url=proxy_url)
        pc: Optional[PooledConnection] = None
        deadline = time.monotonic() + self.acquire_timeout
        discarded: list[PooledConnection] = []

        while pc is None:
            create_new = False
            with self._condition:
                if time.time() - self._last_cleanup > self._cleanup_interval:
                    discarded.extend(self._cleanup_idle_unlocked())

                pool = self._pools.setdefault(pool_key, [])
                for candidate in pool:
                    if not candidate.in_use:
                        candidate.in_use = True
                        pc = candidate
                        break

                if pc is None:
                    creating = self._creating.get(pool_key, 0)
                    if len(pool) + creating < self.max_connections_per_sender:
                        self._creating[pool_key] = creating + 1
                        create_new = True
                    else:
                        remaining = deadline - time.monotonic()
                        if remaining <= 0:
                            raise TimeoutError(f"SMTP connection pool exhausted for {sender.email}")
                        self._condition.wait(timeout=min(0.5, remaining))

            for stale in discarded:
                self._close_connection(stale)
            discarded.clear()

            if create_new:
                try:
                    server = self._create_connection(sender, proxy_url=proxy_url)
                    new_pc = PooledConnection(
                        server=server,
                        sender_email=sender.email,
                        in_use=True,
                        use_count=1,
                    )
                except Exception:
                    with self._condition:
                        self._finish_creation_unlocked(pool_key)
                        self._condition.notify_all()
                    raise
                with self._condition:
                    self._finish_creation_unlocked(pool_key)
                    self._pools.setdefault(pool_key, []).append(new_pc)
                    self._condition.notify_all()
                pc = new_pc
            elif pc is not None:
                try:
                    pc.server.noop()
                    pc.last_used = time.time()
                    pc.use_count += 1
                except Exception:
                    with self._condition:
                        pool = self._pools.get(pool_key, [])
                        if pc in pool:
                            pool.remove(pc)
                        self._condition.notify_all()
                    self._close_connection(pc)
                    pc = None

        try:
            yield pc.server
        except Exception:
            with self._condition:
                pool = self._pools.get(pool_key, [])
                if pc in pool:
                    pool.remove(pc)
                self._condition.notify_all()
            self._close_connection(pc)
            raise
        else:
            with self._condition:
                pc.in_use = False
                pc.last_used = time.time()
                self._condition.notify_all()

    def close_all(self) -> None:
        with self._condition:
            connections = [pc for pool in self._pools.values() for pc in pool]
            self._pools.clear()
            self._creating.clear()
            self._condition.notify_all()
        for pc in connections:
            self._close_connection(pc)


_smtp_pool = SMTPConnectionPool(
    max_connections_per_sender=max(1, int(settings.SMTP_POOL_MAX_PER_SENDER)),
    max_idle_time=max(1, int(settings.SMTP_POOL_IDLE_SECONDS)),
)


# Rate limit / risk-control error patterns
RATE_LIMIT_PATTERNS = [
    "too many attempts",
    "rate limit",
    "spam",
    "blocked",
    "too many",
    "quota",
    "limit exceeded",
    "temporarily deferred",
    "try again later",
    "throttl",
    "421",
    "450",
    "451",
    "452",
    "550 5.7",
]


def is_rate_limit_error(error: str) -> bool:
    if not error:
        return False
    error_lower = error.lower()
    return any(pattern in error_lower for pattern in RATE_LIMIT_PATTERNS)


def is_auth_error(error: str) -> bool:
    if not error:
        return False
    el = error.lower()
    return any(x in el for x in ("auth", "authentication", "login", "535", "534", "invalid credentials"))


class EmailSender:
    MAX_RETRIES = 3
    RETRY_BACKOFF_BASE = 2.0
    SMTP_TIMEOUT = 30
    RATE_LIMIT_CODES = {421, 450, 451, 452}
    MAX_RETRY_DELAY = 3600.0

    def __init__(self):
        self._attachment_cache: OrderedDict[tuple[str, int, int], bytes] = OrderedDict()
        self._attachment_cache_bytes = 0
        self._attachment_cache_limit = max(
            0, int(getattr(settings, "ATTACHMENT_CACHE_MB", 64))
        ) * 1024 * 1024
        self._attachment_cache_lock = threading.Lock()

    def _read_attachment(self, path: str) -> bytes:
        """Read an attachment with a bounded mtime/size keyed LRU cache."""
        stat = os.stat(path)
        key = (path, int(stat.st_mtime_ns), int(stat.st_size))
        if self._attachment_cache_limit > 0 and stat.st_size <= self._attachment_cache_limit:
            with self._attachment_cache_lock:
                cached = self._attachment_cache.get(key)
                if cached is not None:
                    self._attachment_cache.move_to_end(key)
                    return cached
        with open(path, "rb") as stream:
            payload = stream.read()
        if self._attachment_cache_limit <= 0 or len(payload) > self._attachment_cache_limit:
            return payload
        with self._attachment_cache_lock:
            previous = self._attachment_cache.pop(key, None)
            if previous is not None:
                self._attachment_cache_bytes -= len(previous)
            self._attachment_cache[key] = payload
            self._attachment_cache_bytes += len(payload)
            while self._attachment_cache and self._attachment_cache_bytes > self._attachment_cache_limit:
                _, evicted = self._attachment_cache.popitem(last=False)
                self._attachment_cache_bytes -= len(evicted)
        return payload

    def build_message(
        self,
        sender_email: str,
        recipient_email: str,
        recipient_name: str,
        subject: str,
        body_html: str,
        attachments: Optional[List[str]] = None,
        from_name: str = "",
    ) -> MIMEMultipart:
        header_values = (sender_email, recipient_email, recipient_name, subject, from_name)
        if any(
            any(ord(char) < 32 or ord(char) == 127 for char in str(value or ""))
            for value in header_values
        ):
            raise ValueError("Email header contains control characters")
        msg = MIMEMultipart()
        if from_name:
            msg["From"] = formataddr((from_name, sender_email))
        else:
            msg["From"] = sender_email
        msg["To"] = formataddr((recipient_name or "", recipient_email)) if recipient_name else recipient_email
        name = recipient_name or "朋友"
        msg["Subject"] = subject.replace("{name}", name).replace("{email}", recipient_email)
        personalized_body = body_html.replace(
            "{name}", html.escape(name)
        ).replace("{email}", html.escape(recipient_email, quote=True))
        msg.attach(MIMEText(personalized_body, "html", "utf-8"))

        if attachments:
            for path in attachments:
                if not path or not os.path.exists(path):
                    continue
                try:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(self._read_attachment(path))
                    encoders.encode_base64(part)
                    filename = os.path.basename(path)
                    part.add_header("Content-Disposition", "attachment", filename=filename)
                    msg.attach(part)
                except Exception as exc:
                    logger.warning("Attachment skip %s: %s", path, exc)
        return msg

    def send(
        self,
        sender: Sender,
        recipient_email: str,
        recipient_name: str,
        subject: str,
        body_html: str,
        attachments: Optional[List[str]] = None,
        max_retries: Optional[int] = None,
        retry_backoff_base: Optional[float] = None,
        proxy_url: str = "",
    ) -> Tuple[bool, str]:
        if sender.sender_type in ("阿里云邮箱推送", "aliyun_dm", "Aliyun DM"):
            from ..services.aliyun_dm import aliyun_dm_sender

            return aliyun_dm_sender.send(
                sender=sender,
                recipient_email=recipient_email,
                recipient_name=recipient_name,
                subject=subject,
                body_html=body_html,
            )
        return self._send_smtp(
            sender,
            recipient_email,
            recipient_name,
            subject,
            body_html,
            attachments,
            max_retries=max_retries,
            retry_backoff_base=retry_backoff_base,
            proxy_url=proxy_url,
        )

    def _send_smtp(
        self,
        sender: Sender,
        recipient_email: str,
        recipient_name: str,
        subject: str,
        body_html: str,
        attachments: Optional[List[str]] = None,
        max_retries: Optional[int] = None,
        retry_backoff_base: Optional[float] = None,
        proxy_url: str = "",
    ) -> Tuple[bool, str]:
        if max_retries is None:
            max_retries = int(getattr(self, "MAX_RETRIES", 3) or 0)
        if retry_backoff_base is None:
            retry_backoff_base = float(getattr(self, "RETRY_BACKOFF_BASE", 2) or 2)
        base = float(retry_backoff_base or 2)
        from_name = getattr(sender, "aliyun_from_name", "") or ""

        for attempt in range(max_retries + 1):
            try:
                with _smtp_pool.get_connection(sender, proxy_url=proxy_url or "") as server:
                    msg = self.build_message(
                        sender.email,
                        recipient_email,
                        recipient_name,
                        subject,
                        body_html,
                        attachments,
                        from_name=from_name,
                    )
                    server.send_message(msg)
                return True, ""
            except smtplib.SMTPAuthenticationError as e:
                logger.error("SMTP auth failed for %s: %s", sender.email, e)
                return False, f"Auth failed: {e}"
            except smtplib.SMTPResponseException as e:
                error_msg = f"SMTP error (code {e.smtp_code}): {e}"
                if e.smtp_code in self.RATE_LIMIT_CODES or is_rate_limit_error(str(e)):
                    if attempt < max_retries:
                        # Retry jitter is scheduling data, not a security token.
                        backoff = min(
                            self.MAX_RETRY_DELAY,
                            (base ** attempt) + random.uniform(0, 1),  # nosec B311
                        )
                        logger.warning(
                            "Rate limited %s, retry in %.1fs (%s/%s)",
                            sender.email,
                            backoff,
                            attempt + 1,
                            max_retries,
                        )
                        time.sleep(backoff)
                        continue
                    return False, error_msg
                if attempt < max_retries:
                    time.sleep(min(self.MAX_RETRY_DELAY, base ** attempt))
                    continue
                return False, error_msg
            except (smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError, ConnectionError, TimeoutError, OSError) as e:
                if attempt < max_retries:
                    time.sleep(min(self.MAX_RETRY_DELAY, base ** attempt))
                    continue
                return False, f"Connection failed: {e}"
            except Exception as e:
                logger.exception("Send failed for %s: %s", sender.email, e)
                if attempt < max_retries:
                    time.sleep(min(self.MAX_RETRY_DELAY, base ** attempt))
                    continue
                return False, f"Send failed: {e}"
        return False, "Max retries exceeded"

    def test_connection(self, sender: Sender) -> Tuple[bool, str]:
        try:
            with _smtp_pool.get_connection(sender) as server:
                server.noop()
            return True, "SMTP连接成功"
        except Exception as e:
            return False, f"连接失败: {e}"

    def close(self) -> None:
        """Close all pooled SMTP sessions during application shutdown."""
        _smtp_pool.close_all()
        with self._attachment_cache_lock:
            self._attachment_cache.clear()
            self._attachment_cache_bytes = 0


email_sender = EmailSender()
