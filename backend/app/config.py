from pathlib import Path
from pydantic import model_validator
from pydantic_settings import BaseSettings

# Get the project root directory
PROJECT_ROOT = Path(__file__).parent.parent.parent


class Settings(BaseSettings):
    APP_NAME: str = "EmailBulkSending"
    APP_ENV: str = "development"
    DEBUG: bool = True
    ENABLE_DOCS: bool = True

    SECRET_KEY: str = "change-me-to-a-random-secret-key-at-least-32-chars"
    AES_KEY: str = "change-me-to-32-bytes-aes-key-123456"

    JWT_ACCESS_TOKEN_EXPIRE_MINUTES: int = 120
    JWT_REFRESH_TOKEN_EXPIRE_DAYS: int = 7
    JWT_ALGORITHM: str = "HS256"

    DATABASE_URL: str = f"sqlite:///{PROJECT_ROOT / 'backend' / 'data' / 'email_bulk.db'}"
    DB_POOL_SIZE: int = 10
    DB_MAX_OVERFLOW: int = 20
    DB_POOL_TIMEOUT: int = 30
    DB_POOL_RECYCLE: int = 1800

    REDIS_URL: str = "redis://localhost:6379/0"
    CELERY_BROKER_URL: str = "redis://localhost:6379/0"
    CELERY_RESULT_BACKEND: str = "redis://localhost:6379/1"

    UPLOAD_DIR: str = str(PROJECT_ROOT / 'backend' / 'uploads')
    MAX_ATTACHMENT_SIZE_MB: int = 25
    MAX_IMAGE_SIZE_MB: int = 5
    MAX_RECIPIENTS_PER_TASK: int = 100000
    MAX_TASK_BODY_BYTES: int = 5 * 1024 * 1024
    MAX_RECIPIENT_IMPORT_SIZE_MB: int = 25
    MAX_RECIPIENTS_PER_IMPORT: int = 100000

    TRACKING_DOMAIN: str = "http://localhost:8000"

    CORS_ORIGINS: str = "http://localhost:5173,http://localhost:80"

    LOGIN_RATE_LIMIT: str = "5/minute"
    TASK_RATE_LIMIT: str = "10/minute"
    # 开放自助注册（生产环境默认关闭，由管理员在用户管理创建账号）
    ALLOW_REGISTER: bool = True
    REGISTER_RATE_LIMIT: str = "10/hour"
    # ``auto`` uses Redis when available and falls back to the local limiter
    # during development or a Redis outage. ``redis`` makes an outage fail
    # open to the local limiter as well, preserving availability.
    RATE_LIMIT_BACKEND: str = "auto"
    RATE_LIMIT_REDIS_PREFIX: str = "email-bulk:rate-limit"
    RATE_LIMIT_REDIS_TIMEOUT: float = 0.2

    # Send engine / SMTP
    SEND_ENGINE_WORKERS: int = 4
    SMTP_TIMEOUT: int = 30
    SMTP_POOL_MAX_PER_SENDER: int = 5
    SMTP_POOL_IDLE_SECONDS: int = 300
    ATTACHMENT_CACHE_MB: int = 64
    SEND_TASK_LEASE_SECONDS: int = 120
    MAX_TASK_CONCURRENCY: int = 64
    MAX_GLOBAL_SEND_CONCURRENCY: int = 64

    # Risk control defaults
    DEFAULT_RATE_LIMIT_COOLDOWN: int = 60
    DEFAULT_RISK_PAUSE_SECONDS: int = 300
    DEFAULT_MAX_CONSECUTIVE_RATE_LIMITS: int = 5
    ALIYUN_MAX_RETRIES: int = 3
    ALIYUN_RETRY_BACKOFF_BASE: float = 1.0
    CELERY_TASK_WAIT_SECONDS: int = 86400

    # 首次启动自动创建的管理员账号（生产环境务必通过 .env 修改密码）
    DEFAULT_ADMIN_USERNAME: str = "admin"
    DEFAULT_ADMIN_PASSWORD: str = "admin123"
    DEFAULT_ADMIN_EMAIL: str = "admin@example.com"

    model_config = {"env_file": ".env", "env_file_encoding": "utf-8", "extra": "ignore"}

    @model_validator(mode="after")
    def validate_production_secrets(self):
        if self.APP_ENV.strip().lower() in {"production", "prod"}:
            secret_markers = ("change-me", "change-in-production", "dev-secret", "dev-aes", "default")
            insecure = (
                len(self.SECRET_KEY) < 32
                or len(self.AES_KEY) < 32
                or any(marker in self.SECRET_KEY.lower() for marker in secret_markers)
                or any(marker in self.AES_KEY.lower() for marker in secret_markers)
            )
            if insecure:
                raise ValueError(
                    "Secure production secrets are required for SECRET_KEY and AES_KEY"
                )
        return self

    @property
    def cors_origins_list(self) -> list[str]:
        return [origin.strip() for origin in self.CORS_ORIGINS.split(",") if origin.strip()]


settings = Settings()
