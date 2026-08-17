import time
from datetime import timedelta
from typing import Optional
from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives import padding as crypto_padding
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from jose import JWTError, jwt
from passlib.context import CryptContext
import base64
import secrets
import hashlib
import hmac
import re
import threading
from ..config import settings
from .time import utcnow

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

# Rate limiting configuration
RATE_LIMIT_WINDOW = 60  # seconds
RATE_LIMIT_MAX_REQUESTS = 100  # requests per window
RATE_LIMIT_MAX_KEYS = 10000
_rate_limit_store = {}
_rate_limit_lock = threading.Lock()
_redis_client = None
_redis_lock = threading.Lock()
_redis_unavailable_until = 0.0
_REDIS_RETRY_SECONDS = 30.0


def _get_redis_client():
    """Return a short-timeout Redis client, or ``None`` when unavailable.

    Rate limiting must never make the authentication endpoint hang when Redis
    is down. The client is lazy and failures are cached briefly to avoid a
    connection attempt for every request during an outage.
    """
    global _redis_client, _redis_unavailable_until
    backend = str(getattr(settings, "RATE_LIMIT_BACKEND", "auto") or "auto").lower()
    if backend in {"memory", "local", "disabled"}:
        return None
    if backend == "auto" and str(getattr(settings, "APP_ENV", "development")).lower() not in {"production", "prod"}:
        # Development and test databases are frequently shared with a local
        # Redis instance. Keep them deterministic; production can opt in via
        # APP_ENV=production, while explicit ``redis`` always enables it.
        return None
    now = time.monotonic()
    if now < _redis_unavailable_until:
        return None
    with _redis_lock:
        if time.monotonic() < _redis_unavailable_until:
            return None
        if _redis_client is not None:
            return _redis_client
        try:
            import redis  # type: ignore

            client = redis.Redis.from_url(
                settings.REDIS_URL,
                socket_connect_timeout=float(getattr(settings, "RATE_LIMIT_REDIS_TIMEOUT", 0.2)),
                socket_timeout=float(getattr(settings, "RATE_LIMIT_REDIS_TIMEOUT", 0.2)),
                decode_responses=True,
            )
            # A bounded ping confirms configuration without blocking callers.
            client.ping()
            _redis_client = client
            return client
        except Exception:
            _redis_client = None
            _redis_unavailable_until = time.monotonic() + _REDIS_RETRY_SECONDS
            return None


def _redis_bucket_key(key: str, window: int, now: Optional[float] = None) -> str:
    now = time.time() if now is None else now
    bucket = int(now // max(1, int(window)))
    digest = hashlib.sha256(str(key).encode("utf-8")).hexdigest()
    prefix = str(getattr(settings, "RATE_LIMIT_REDIS_PREFIX", "email-bulk:rate-limit"))
    return f"{prefix}:{int(window)}:{bucket}:{digest}"


def _check_rate_limit_redis(key: str, max_requests: int, window: int) -> Optional[bool]:
    client = _get_redis_client()
    if client is None:
        return None
    redis_key = _redis_bucket_key(key, window)
    try:
        count = int(client.incr(redis_key))
        if count == 1:
            client.expire(redis_key, max(1, int(window)) + 2)
        return count <= max_requests
    except Exception:
        # Mark the client unavailable and transparently use the local fallback.
        global _redis_client, _redis_unavailable_until
        with _redis_lock:
            _redis_client = None
            _redis_unavailable_until = time.monotonic() + _REDIS_RETRY_SECONDS
        return None


def _get_rate_limit_remaining_redis(key: str, max_requests: int, window: int) -> Optional[int]:
    client = _get_redis_client()
    if client is None:
        return None
    try:
        value = client.get(_redis_bucket_key(key, window))
        return max(0, max_requests - int(value or 0))
    except Exception:
        global _redis_client, _redis_unavailable_until
        with _redis_lock:
            _redis_client = None
            _redis_unavailable_until = time.monotonic() + _REDIS_RETRY_SECONDS
        return None

def verify_password(plain_password: str, hashed_password: str) -> bool:
    try:
        return pwd_context.verify(plain_password, hashed_password)
    except Exception:
        # Malformed legacy hashes should behave like invalid credentials, not
        # turn a login attempt into a 500 response.
        return False

def get_password_hash(password: str) -> str:
    return pwd_context.hash(password)

def create_access_token(data: dict, expires_delta: Optional[timedelta] = None) -> str:
    to_encode = data.copy()
    to_encode["sub"] = str(to_encode["sub"])
    if expires_delta:
        expire = utcnow() + expires_delta
    else:
        expire = utcnow() + timedelta(minutes=settings.JWT_ACCESS_TOKEN_EXPIRE_MINUTES)
    to_encode.update({"exp": expire, "type": "access"})
    encoded_jwt = jwt.encode(to_encode, settings.SECRET_KEY, algorithm=settings.JWT_ALGORITHM)
    return encoded_jwt

def create_refresh_token(data: dict) -> str:
    to_encode = data.copy()
    to_encode["sub"] = str(to_encode["sub"])
    expire = utcnow() + timedelta(days=settings.JWT_REFRESH_TOKEN_EXPIRE_DAYS)
    to_encode.update({"exp": expire, "type": "refresh"})
    encoded_jwt = jwt.encode(to_encode, settings.SECRET_KEY, algorithm=settings.JWT_ALGORITHM)
    return encoded_jwt

def decode_token(token: str) -> Optional[dict]:
    try:
        payload = jwt.decode(token, settings.SECRET_KEY, algorithms=[settings.JWT_ALGORITHM])
        subject = payload.get("sub")
        if subject is None or str(subject).strip() == "":
            return None
        payload["sub"] = int(subject)
        if payload["sub"] <= 0:
            return None
        return payload
    except (JWTError, ValueError, TypeError, KeyError):
        return None

# AES encryption with random IV (secure)
def encrypt_password(password: str) -> str:
    """Encrypt a secret with authenticated AES-GCM (v3 format)."""
    key = hashlib.sha256(settings.AES_KEY.encode("utf-8")).digest()
    nonce = secrets.token_bytes(12)
    encrypted = AESGCM(key).encrypt(nonce, password.encode("utf-8"), None)
    ciphertext, tag = encrypted[:-16], encrypted[-16:]
    return "v3:" + base64.b64encode(nonce + tag + ciphertext).decode("ascii")


def _decrypt_legacy_cbc(key: bytes, iv: bytes, ciphertext: bytes) -> str:
    """Decrypt historical CBC records without extending their use for writes."""
    if len(key) not in {16, 24, 32} or len(iv) != 16:
        raise ValueError("Invalid encrypted secret")
    if not ciphertext or len(ciphertext) % 16:
        raise ValueError("Invalid encrypted secret")
    try:
        decryptor = Cipher(algorithms.AES(key), modes.CBC(iv)).decryptor()
        padded = decryptor.update(ciphertext) + decryptor.finalize()
        unpadder = crypto_padding.PKCS7(algorithms.AES.block_size).unpadder()
        plaintext = unpadder.update(padded) + unpadder.finalize()
        return plaintext.decode("utf-8")
    except (ValueError, UnicodeDecodeError) as exc:
        raise ValueError("Invalid encrypted secret") from exc

def decrypt_password(encrypted: str) -> str:
    """Decrypt v3 secrets, with read compatibility for legacy v2/v1 data."""
    if not isinstance(encrypted, str) or not encrypted:
        raise ValueError("Invalid encrypted secret")
    if encrypted.startswith("v3:"):
        try:
            data = base64.b64decode(encrypted[3:].encode("ascii"), validate=True)
            if len(data) < 28:
                raise ValueError("truncated payload")
            nonce, tag, ciphertext = data[:12], data[12:28], data[28:]
            key = hashlib.sha256(settings.AES_KEY.encode("utf-8")).digest()
            plaintext = AESGCM(key).decrypt(nonce, ciphertext + tag, None)
            return plaintext.decode("utf-8")
        except (InvalidTag, ValueError, UnicodeDecodeError, base64.binascii.Error) as exc:
            raise ValueError("Invalid encrypted secret") from exc
    if encrypted.startswith("v2:"):
        try:
            key = settings.AES_KEY.encode("utf-8")[:32]
            data = base64.b64decode(encrypted[3:].encode("ascii"), validate=True)
            if len(data) < 32:
                raise ValueError("Invalid encrypted secret")
            return _decrypt_legacy_cbc(key, data[:16], data[16:])
        except (ValueError, UnicodeDecodeError, base64.binascii.Error) as exc:
            raise ValueError("Invalid encrypted secret") from exc
    return decrypt_password_v1(encrypted)

def decrypt_password_v1(encrypted: str) -> str:
    """Decrypt legacy v1 format (fixed IV)."""
    try:
        key = settings.AES_KEY.encode("utf-8")[:32]
        ciphertext = base64.b64decode(encrypted.encode("ascii"), validate=True)
        return _decrypt_legacy_cbc(key, key[:16], ciphertext)
    except (ValueError, UnicodeDecodeError, base64.binascii.Error) as exc:
        raise ValueError("Invalid encrypted secret") from exc

# Backward compatibility alias
def decrypt_password_v2(encrypted: str) -> str:
    """Alias for decrypt_password for backward compatibility."""
    return decrypt_password(encrypted)

# Rate limiting functions
def check_rate_limit(key: str, max_requests: int = RATE_LIMIT_MAX_REQUESTS, window: int = RATE_LIMIT_WINDOW) -> bool:
    """Check if request is within rate limit. Returns True if allowed."""
    max_requests = max(1, int(max_requests))
    window = max(1, int(window))
    distributed = _check_rate_limit_redis(key, max_requests, window)
    if distributed is not None:
        return distributed
    now = time.time()
    with _rate_limit_lock:
        entries = _rate_limit_store.get(key, [])
        entries = [timestamp for timestamp in entries if now - timestamp < window]
        if len(entries) >= max_requests:
            _rate_limit_store[key] = entries
            return False
        entries.append(now)
        _rate_limit_store[key] = entries
        if len(_rate_limit_store) > RATE_LIMIT_MAX_KEYS:
            # Bound memory under an IP/username spray attack. Expired entries
            # are removed first; if all are fresh, evict the oldest key.
            cutoff = now - window
            stale_keys = [
                stored_key
                for stored_key, timestamps in _rate_limit_store.items()
                if not timestamps or timestamps[-1] < cutoff
            ]
            for stored_key in stale_keys[: max(1, len(stale_keys) - RATE_LIMIT_MAX_KEYS // 2)]:
                _rate_limit_store.pop(stored_key, None)
            while len(_rate_limit_store) > RATE_LIMIT_MAX_KEYS:
                oldest_key = min(
                    _rate_limit_store,
                    key=lambda stored_key: _rate_limit_store[stored_key][-1]
                    if _rate_limit_store[stored_key]
                    else 0,
                )
                _rate_limit_store.pop(oldest_key, None)
        return True

def get_rate_limit_remaining(key: str, max_requests: int = RATE_LIMIT_MAX_REQUESTS, window: int = RATE_LIMIT_WINDOW) -> int:
    """Get remaining requests in current window."""
    max_requests = max(1, int(max_requests))
    window = max(1, int(window))
    distributed = _get_rate_limit_remaining_redis(key, max_requests, window)
    if distributed is not None:
        return distributed
    now = time.time()
    with _rate_limit_lock:
        entries = _rate_limit_store.get(key, [])
        entries = [timestamp for timestamp in entries if now - timestamp < window]
        _rate_limit_store[key] = entries
        return max(0, max_requests - len(entries))


def parse_rate_limit(rule: str) -> tuple[int, int]:
    """Parse rules such as ``5/minute`` into a request count and window."""
    match = re.fullmatch(
        r"\s*(\d+)\s*/\s*(second|minute|hour|day)s?\s*",
        str(rule or ""),
        flags=re.IGNORECASE,
    )
    if not match:
        raise ValueError(f"Invalid rate limit rule: {rule!r}")
    count = max(1, int(match.group(1)))
    seconds = {"second": 1, "minute": 60, "hour": 3600, "day": 86400}
    return count, seconds[match.group(2).lower()]


def check_configured_rate_limit(key: str, rule: str) -> bool:
    max_requests, window = parse_rate_limit(rule)
    return check_rate_limit(key, max_requests=max_requests, window=window)


def create_tracking_signature(log_id: int, action: str) -> str:
    payload = f"{action}:{int(log_id)}".encode("utf-8")
    return hmac.new(
        settings.SECRET_KEY.encode("utf-8"), payload, hashlib.sha256
    ).hexdigest()


def verify_tracking_signature(log_id: int, action: str, signature: str) -> bool:
    if not signature:
        return False
    expected = create_tracking_signature(log_id, action)
    return hmac.compare_digest(expected, str(signature))
