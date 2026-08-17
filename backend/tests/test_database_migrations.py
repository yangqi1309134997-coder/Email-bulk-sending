from sqlalchemy import create_engine, inspect, text

from app.database import migrate_database


def test_migrate_database_upgrades_legacy_schema_without_losing_rows(tmp_path):
    database_url = f"sqlite:///{tmp_path / 'legacy.db'}"
    engine = create_engine(database_url)

    with engine.begin() as connection:
        connection.execute(text("""
            CREATE TABLE senders (
                id INTEGER PRIMARY KEY,
                email VARCHAR NOT NULL,
                smtp_port INTEGER NOT NULL,
                use_tls BOOLEAN NOT NULL
            )
        """))
        connection.execute(text("""
            CREATE TABLE send_tasks (
                id INTEGER PRIMARY KEY,
                status VARCHAR NOT NULL
            )
        """))
        connection.execute(text("""
            CREATE TABLE send_logs (
                id INTEGER PRIMARY KEY,
                status VARCHAR NOT NULL
            )
        """))
        connection.execute(
            text("INSERT INTO senders (id, email, smtp_port, use_tls) VALUES (1, 'ssl@example.com', 465, 1)")
        )
        connection.execute(
            text("INSERT INTO senders (id, email, smtp_port, use_tls) VALUES (2, 'tls@example.com', 587, 1)")
        )

    migrate_database(engine)
    migrate_database(engine)

    inspector = inspect(engine)
    sender_columns = {column["name"] for column in inspector.get_columns("senders")}
    task_columns = {column["name"] for column in inspector.get_columns("send_tasks")}
    log_columns = {column["name"] for column in inspector.get_columns("send_logs")}

    assert {
        "smtp_security",
        "smtp_username",
        "cb_state",
        "cb_failure_count",
        "cb_success_count",
        "cb_next_attempt_time",
        "cb_last_failure_time",
    } <= sender_columns
    assert {"next_run_at", "pause_reason", "lease_owner", "lease_expires_at", "last_heartbeat_at"} <= task_columns
    assert {"attempt_count", "claimed_at", "next_attempt_at", "last_error_code"} <= log_columns

    with engine.connect() as connection:
        rows = connection.execute(
            text("SELECT id, smtp_security FROM senders ORDER BY id")
        ).all()
    assert rows == [(1, "ssl"), (2, "starttls")]

