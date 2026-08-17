from types import SimpleNamespace

from app.tasks import scheduled_tasks


class FakeResult:
    def all(self):
        return []


class FakeSession:
    def __init__(self, database_engine):
        assert database_engine == "database-engine"

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        return False

    def exec(self, statement):
        return FakeResult()


def test_recover_stuck_tasks_uses_database_engine_not_send_engine(monkeypatch):
    fake_send_engine = SimpleNamespace(submit=lambda task_id: None)
    monkeypatch.setattr(scheduled_tasks, "engine", "database-engine")
    monkeypatch.setattr(scheduled_tasks, "Session", FakeSession)
    monkeypatch.setattr(scheduled_tasks, "get_send_engine", lambda: fake_send_engine)

    result = scheduled_tasks.recover_stuck_tasks.run()

    assert result == {"recovered": 0}


def test_manual_pause_is_never_auto_resumed():
    task = SimpleNamespace(
        status="paused",
        pause_reason="manual",
        smart_config='{"auto_resume_after_cooldown": true}',
        next_run_at=None,
    )

    assert scheduled_tasks.should_auto_resume_task(task) is False


def test_string_false_auto_resume_is_respected():
    task = SimpleNamespace(
        status="paused",
        pause_reason="rate_limit",
        smart_config='{"auto_resume_after_cooldown": "false"}',
        next_run_at=None,
    )

    assert scheduled_tasks.should_auto_resume_task(task) is False


def test_sender_id_json_parser_preserves_valid_unique_ids():
    assert scheduled_tasks.parse_sender_ids('[3, "2", 3, -1]') == [3, 2]
    assert scheduled_tasks.parse_sender_ids('[true, 1.5, "bad", 4, " 2 "]') == [4, 2]
    assert scheduled_tasks.parse_sender_ids('[3, "invalid", 2]') == [3, 2]
    assert scheduled_tasks.parse_sender_ids('{"id": 3}') == []
    assert scheduled_tasks.parse_sender_ids("invalid") == []
