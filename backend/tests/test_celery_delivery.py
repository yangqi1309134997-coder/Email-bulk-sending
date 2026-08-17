from types import SimpleNamespace

from app.tasks import celery_app as celery_module
from app.tasks import send_email


class _Task:
    def __init__(self, statuses):
        self.statuses = iter(statuses)
        self.status = None


class _Session:
    statuses = iter(["running", "completed"])

    def __init__(self, engine):
        self.task = _Task([])

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def get(self, model, task_id):
        self.task.status = next(type(self).statuses, "completed")
        return self.task


def test_celery_task_waits_for_persistent_engine_completion(monkeypatch):
    submitted = []
    monkeypatch.setattr(send_email, "Session", _Session)
    monkeypatch.setattr(send_email, "engine", "db")
    monkeypatch.setattr(
        send_email,
        "get_send_engine",
        lambda: SimpleNamespace(submit=lambda task_id: submitted.append(task_id)),
    )
    monkeypatch.setattr(send_email.time, "sleep", lambda _: None)

    result = send_email.send_batch_task.run(42)

    assert submitted == [42]
    assert result == {"task_id": 42, "status": "completed"}


def test_celery_uses_late_ack_and_single_prefetch():
    config = celery_module.celery_app.conf
    assert config.task_acks_late is True
    assert config.task_reject_on_worker_lost is True
    assert config.worker_prefetch_multiplier == 1
