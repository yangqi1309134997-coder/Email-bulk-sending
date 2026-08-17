"""Database engine, session helpers, and lightweight schema migrations."""

from __future__ import annotations

import logging
import sqlite3
from datetime import date, datetime

from sqlalchemy import and_, column, event, inspect, or_, table as sql_table, text, update
from sqlmodel import SQLModel, Session, create_engine, select

from .config import settings

logger = logging.getLogger(__name__)

_engine_options = {
    "echo": settings.DEBUG,
    "pool_pre_ping": True,
}
if settings.DATABASE_URL.startswith("sqlite"):
    # Python 3.12 deprecated sqlite3's implicit date/datetime adapters. Keep
    # the historical ISO storage format explicit so future Python versions do
    # not break timestamp writes when the defaults are removed.
    sqlite3.register_adapter(date, lambda value: value.isoformat())
    sqlite3.register_adapter(datetime, lambda value: value.isoformat(" "))
    _engine_options["connect_args"] = {"check_same_thread": False}
else:
    _engine_options.update(
        pool_size=max(1, int(settings.DB_POOL_SIZE)),
        max_overflow=max(0, int(settings.DB_MAX_OVERFLOW)),
        pool_timeout=max(1, int(settings.DB_POOL_TIMEOUT)),
        pool_recycle=max(60, int(settings.DB_POOL_RECYCLE)),
    )

engine = create_engine(settings.DATABASE_URL, **_engine_options)


if settings.DATABASE_URL.startswith("sqlite"):
    @event.listens_for(engine, "connect")
    def _configure_sqlite(dbapi_connection, connection_record) -> None:
        cursor = dbapi_connection.cursor()
        cursor.execute("PRAGMA foreign_keys=ON")
        cursor.execute("PRAGMA busy_timeout=10000")
        cursor.execute("PRAGMA journal_mode=WAL")
        cursor.execute("PRAGMA synchronous=NORMAL")
        cursor.close()


def create_db_and_tables() -> None:
    # Import models so SQLModel metadata is fully populated
    from . import models  # noqa: F401

    SQLModel.metadata.create_all(engine)
    migrate_database(engine)


def get_session():
    with Session(engine) as session:
        yield session


def ensure_default_admin() -> None:
    """Create the default administrator on first boot (idempotent).

    Runs on every startup so a fresh deployment (or an empty PostgreSQL
    volume) always has a usable login. Existing databases are untouched.
    """
    from .models.user import User
    from .utils.security import get_password_hash
    from .config import settings

    with Session(engine) as session:
        existing = session.exec(select(User).limit(1)).first()
        if existing is not None:
            return
        admin = User(
            username=str(settings.DEFAULT_ADMIN_USERNAME).strip(),
            password_hash=get_password_hash(str(settings.DEFAULT_ADMIN_PASSWORD)),
            email=str(settings.DEFAULT_ADMIN_EMAIL).strip(),
            role="admin",
        )
        session.add(admin)
        session.commit()
        logger.warning(
            "Created default admin account '%s' (change the password in production)",
            admin.username,
        )


def _existing_columns(connection, table: str) -> set[str]:
    try:
        return {column["name"] for column in inspect(connection).get_columns(table)}
    except Exception:
        return set()


def _table_exists(connection, table: str) -> bool:
    try:
        return bool(inspect(connection).has_table(table))
    except Exception:
        return False


def _quote_identifier(connection, identifier: str) -> str:
    """Quote a known schema identifier for every supported SQL dialect."""
    return connection.dialect.identifier_preparer.quote(identifier)


def _add_column(connection, table: str, name: str, ddl: str) -> None:
    table_sql = _quote_identifier(connection, table)
    name_sql = _quote_identifier(connection, name)
    # PostgreSQL supports IF NOT EXISTS and avoids startup races between
    # multiple API workers. Other dialects are guarded by introspection below.
    if connection.dialect.name == "postgresql":
        connection.execute(
            text(f"ALTER TABLE {table_sql} ADD COLUMN IF NOT EXISTS {name_sql} {ddl}")
        )
    else:
        connection.execute(text(f"ALTER TABLE {table_sql} ADD COLUMN {name_sql} {ddl}"))


def _dialect_ddl(dialect: str, ddl: str) -> str:
    """Return portable type/default syntax for additive column migrations."""
    if dialect == "sqlite":
        return ddl
    # DATETIME is accepted by MySQL but not PostgreSQL. TIMESTAMP is accepted
    # by all supported engines and these columns are intentionally nullable.
    ddl = ddl.replace("DATETIME", "TIMESTAMP")
    ddl = ddl.replace("BOOLEAN DEFAULT 1", "BOOLEAN DEFAULT TRUE")
    if dialect in {"mysql", "mariadb"}:
        # MySQL versions commonly used in deployments reject defaults on TEXT
        # columns. They remain nullable for legacy rows and are normalized by
        # application defaults when loaded.
        ddl = ddl.replace("TEXT DEFAULT ''", "TEXT")
        ddl = ddl.replace("TEXT DEFAULT '[]'", "TEXT")
        ddl = ddl.replace("TEXT DEFAULT '{}'", "TEXT")
    return ddl


def _ensure_index(connection, table: str, index_name: str, columns: list[str]) -> None:
    if not _table_exists(connection, table):
        return
    try:
        existing = {item.get("name") for item in inspect(connection).get_indexes(table)}
        if index_name in existing:
            return
        table_sql = _quote_identifier(connection, table)
        index_sql = _quote_identifier(connection, index_name)
        columns_sql = ", ".join(_quote_identifier(connection, column) for column in columns)
        connection.execute(
            text(f"CREATE INDEX {index_sql} ON {table_sql} ({columns_sql})")
        )
    except Exception as exc:
        # A concurrent worker may have created the index after introspection.
        # Do not make application startup fail in that benign race.
        logger.debug("Unable to create index %s: %s", index_name, exc)


def migrate_database(db_engine=None) -> None:
    """Apply idempotent additive migrations on all supported SQL dialects.

    Safe to call repeatedly. Adds columns required by current models without
    dropping data. Also backfills simple derived values where possible. This is
    intentionally a small bootstrap migration for deployments that do not yet
    run Alembic; production installations can still use Alembic afterwards.
    """
    db_engine = db_engine or engine
    dialect = db_engine.dialect.name

    sender_columns = [
        ("user_id", "INTEGER"),
        ("email", "VARCHAR"),
        ("password", "VARCHAR DEFAULT ''"),
        ("smtp_server", "VARCHAR DEFAULT ''"),
        ("smtp_port", "INTEGER DEFAULT 587"),
        ("use_tls", "BOOLEAN DEFAULT 1"),
        ("sender_type", "VARCHAR DEFAULT '自定义SMTP'"),
        ("enabled", "BOOLEAN DEFAULT 1"),
        ("weight", "INTEGER DEFAULT 50"),
        ("daily_quota", "INTEGER DEFAULT 500"),
        ("daily_sent", "INTEGER DEFAULT 0"),
        ("success_rate", "FLOAT DEFAULT 1.0"),
        ("status", "VARCHAR DEFAULT 'active'"),
        ("consecutive_failures", "INTEGER DEFAULT 0"),
        ("paused_until", "DATETIME"),
        ("aliyun_access_key", "TEXT DEFAULT ''"),
        ("aliyun_access_secret", "TEXT DEFAULT ''"),
        ("aliyun_region", "TEXT DEFAULT 'cn-hangzhou'"),
        ("aliyun_from_name", "TEXT DEFAULT ''"),
        ("created_at", "DATETIME"),
        ("updated_at", "DATETIME"),
        ("cb_state", "TEXT DEFAULT 'closed'"),
        ("cb_failure_count", "INTEGER DEFAULT 0"),
        ("cb_success_count", "INTEGER DEFAULT 0"),
        ("cb_next_attempt_time", "DATETIME"),
        ("cb_last_failure_time", "DATETIME"),
        # Optional compatibility / future fields
        ("smtp_username", "VARCHAR DEFAULT ''"),
        ("smtp_security", "VARCHAR DEFAULT ''"),
    ]

    task_columns = [
        ("user_id", "INTEGER"),
        ("name", "VARCHAR DEFAULT ''"),
        ("status", "VARCHAR DEFAULT 'pending'"),
        ("sender_ids", "TEXT DEFAULT '[]'"),
        ("recipient_count", "INTEGER DEFAULT 0"),
        ("success_count", "INTEGER DEFAULT 0"),
        ("fail_count", "INTEGER DEFAULT 0"),
        ("open_count", "INTEGER DEFAULT 0"),
        ("click_count", "INTEGER DEFAULT 0"),
        ("subject", "TEXT DEFAULT ''"),
        ("body", "TEXT DEFAULT ''"),
        ("attachments", "TEXT DEFAULT '[]'"),
        ("schedule_type", "VARCHAR DEFAULT 'immediate'"),
        ("schedule_time", "DATETIME"),
        ("smart_config", "TEXT DEFAULT '{}'"),
        ("delay_min", "INTEGER DEFAULT 5"),
        ("delay_max", "INTEGER DEFAULT 15"),
        ("proxies", "TEXT DEFAULT '[]'"),
        ("load_balance_strategy", "VARCHAR DEFAULT 'round_robin'"),
        ("created_at", "DATETIME"),
        ("completed_at", "DATETIME"),
        # Optional scheduler/lease fields for resilience
        ("next_run_at", "DATETIME"),
        ("pause_reason", "VARCHAR DEFAULT ''"),
        ("lease_owner", "VARCHAR DEFAULT ''"),
        ("lease_expires_at", "DATETIME"),
        ("last_heartbeat_at", "DATETIME"),
    ]

    log_columns = [
        ("task_id", "INTEGER"),
        ("sender_id", "INTEGER"),
        ("recipient_email", "VARCHAR DEFAULT ''"),
        ("recipient_name", "VARCHAR DEFAULT ''"),
        ("subject", "TEXT DEFAULT ''"),
        ("status", "VARCHAR DEFAULT 'pending'"),
        ("error_message", "TEXT DEFAULT ''"),
        ("sent_at", "DATETIME"),
        ("opened_at", "DATETIME"),
        ("clicked_at", "DATETIME"),
        ("attempt_count", "INTEGER DEFAULT 0"),
        ("claimed_at", "DATETIME"),
        ("next_attempt_at", "DATETIME"),
        ("last_error_code", "VARCHAR DEFAULT ''"),
    ]

    template_columns = [
        ("user_id", "INTEGER"),
        ("name", "VARCHAR DEFAULT ''"),
        ("description", "TEXT DEFAULT ''"),
        ("sender_type", "VARCHAR DEFAULT '自定义SMTP'"),
        ("smtp_server", "VARCHAR DEFAULT ''"),
        ("smtp_port", "INTEGER DEFAULT 587"),
        ("use_tls", "BOOLEAN DEFAULT 1"),
        ("smtp_username", "VARCHAR DEFAULT ''"),
        ("smtp_security", "VARCHAR DEFAULT ''"),
        ("weight", "INTEGER DEFAULT 50"),
        ("daily_quota", "INTEGER DEFAULT 500"),
        ("aliyun_access_key", "TEXT DEFAULT ''"),
        ("aliyun_access_secret", "TEXT DEFAULT ''"),
        ("aliyun_region", "TEXT DEFAULT 'cn-hangzhou'"),
        ("aliyun_from_name", "TEXT DEFAULT ''"),
        ("created_at", "DATETIME"),
        ("updated_at", "DATETIME"),
    ]

    user_columns = [
        ("username", "VARCHAR DEFAULT ''"),
        ("password_hash", "TEXT DEFAULT ''"),
        ("role", "VARCHAR DEFAULT 'operator'"),
        ("email", "VARCHAR DEFAULT ''"),
        ("created_at", "DATETIME"),
        ("last_login", "DATETIME"),
    ]

    content_template_columns = [
        ("user_id", "INTEGER"),
        ("name", "VARCHAR DEFAULT ''"),
        ("subject", "TEXT DEFAULT ''"),
        ("body", "TEXT DEFAULT ''"),
        ("variables", "TEXT DEFAULT '[]'"),
        ("created_at", "DATETIME"),
        ("updated_at", "DATETIME"),
    ]

    tables = {
        "senders": sender_columns,
        "send_tasks": task_columns,
        "send_logs": log_columns,
        "sender_templates": template_columns,
        "user": user_columns,
        "templates": content_template_columns,
    }

    with db_engine.begin() as connection:
        # PostgreSQL advisory locks serialize startup migrations across API
        # workers without requiring a separate migration service.
        if dialect == "postgresql":
            connection.execute(
                text("SELECT pg_advisory_xact_lock(hashtext('email_bulk_schema_v1'))")
            )

        for table, additions in tables.items():
            if not _table_exists(connection, table):
                continue
            existing = _existing_columns(connection, table)
            for name, ddl in additions:
                if name in existing:
                    continue
                try:
                    _add_column(connection, table, name, _dialect_ddl(dialect, ddl))
                    existing.add(name)
                except Exception as exc:
                    # A concurrent worker may have added the same column. Check
                    # again before surfacing a genuine migration failure.
                    if name not in _existing_columns(connection, table):
                        logger.warning("Unable to add %s.%s: %s", table, name, exc)

        # Backfill SMTP security for legacy sender rows. Use quoted names since
        # this code also runs on databases where identifiers may be reserved.
        if _table_exists(connection, "senders"):
            existing = _existing_columns(connection, "senders")
            if {"smtp_security", "smtp_port"} <= existing:
                senders_table = sql_table(
                    "senders",
                    column("smtp_security"),
                    column("smtp_port"),
                )
                security_col = senders_table.c.smtp_security
                port_col = senders_table.c.smtp_port
                connection.execute(
                    update(senders_table)
                    .where(
                        and_(
                            or_(security_col.is_(None), security_col == ""),
                            port_col == 465,
                        )
                    )
                    .values(smtp_security="ssl")
                )
                connection.execute(
                    update(senders_table)
                    .where(
                        and_(
                            or_(security_col.is_(None), security_col == ""),
                            port_col != 465,
                            port_col.is_not(None),
                        )
                    )
                    .values(smtp_security="starttls")
                )

        _ensure_index(connection, "send_tasks", "idx_send_tasks_due", ["status", "next_run_at"])
        _ensure_index(connection, "send_tasks", "idx_send_tasks_lease", ["lease_expires_at"])
        _ensure_index(
            connection,
            "send_tasks",
            "idx_send_tasks_user_created",
            ["user_id", "created_at"],
        )
        _ensure_index(
            connection,
            "senders",
            "idx_senders_user_email",
            ["user_id", "email"],
        )
        _ensure_index(
            connection,
            "send_logs",
            "idx_send_logs_pending",
            ["task_id", "status", "next_attempt_at"],
        )
        _ensure_index(
            connection,
            "templates",
            "idx_templates_user_updated",
            ["user_id", "updated_at"],
        )
