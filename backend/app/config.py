import os
from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    APP_NAME: str = "EmailBulkSending"
    APP_ENV: str = "development"
    DEBUG: bool = True

    SECRET_KEY: str = "change-me-to-a-random-secret-key-at-least-32-chars"
    AES_KEY: str = "change-me-to-32-bytes-aes-key-123456"

    JWT_ACCESS_TOKEN_EXPIRE_MINUTES: int = 120
    JWT_REFRESH_TOKEN_EXPIRE_DAYS: int = 7
    JWT_ALGORITHM: str = "HS256"

    DATABASE_URL: str = "sqlite:///./data/email_bulk.db"

    REDIS_URL: str = "redis://redis:6379/0"
    CELERY_BROKER_URL: str = "redis://redis:6379/0"
    CELERY_RESULT_BACKEND: str = "redis://redis:6379/1"

    UPLOAD_DIR: str = "./uploads"
    MAX_ATTACHMENT_SIZE_MB: int = 25
    MAX_IMAGE_SIZE_MB: int = 5

    TRACKING_DOMAIN: str = "http://localhost:8000"

    CORS_ORIGINS: str = "http://localhost:5173,http://localhost:80"

    LOGIN_RATE_LIMIT: str = "5/minute"
    TASK_RATE_LIMIT: str = "10/minute"

    model_config = {"env_file": ".env", "env_file_encoding": "utf-8", "extra": "ignore"}

    @property
    def cors_origins_list(self) -> list[str]:
        return [origin.strip() for origin in self.CORS_ORIGINS.split(",") if origin.strip()]


settings = Settings()
