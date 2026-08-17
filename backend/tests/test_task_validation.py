import pytest
from pydantic import ValidationError

from app.api.tasks import TaskCreate


def _payload(smart_config):
    return {
        "name": "validated task",
        "sender_ids": [1],
        "subject": "subject",
        "body": "<p>body</p>",
        "recipients": [{"email": "recipient@example.com", "name": "Recipient"}],
        "smart_config": smart_config,
    }


def test_camel_case_boolean_false_is_not_coerced_to_true():
    task = TaskCreate.model_validate(
        _payload(
            {
                "autoResumeAfterCooldown": "false",
                "concurrencyPerSender": 4,
                "batchSize": 100,
            }
        )
    )

    assert task.smart_config.auto_resume_after_cooldown is False


def test_invalid_proxy_path_and_header_controls_are_rejected():
    payload = _payload({})
    payload["proxies"] = ["http://proxy.example.com:8080/path"]
    with pytest.raises(ValidationError):
        TaskCreate.model_validate(payload)

    payload = _payload({})
    payload["subject"] = "subject\r\nBcc: attacker@example.com"
    with pytest.raises(ValidationError):
        TaskCreate.model_validate(payload)

    payload = _payload({})
    payload["recipients"][0]["name"] = "name\nforged"
    with pytest.raises(ValidationError):
        TaskCreate.model_validate(payload)


@pytest.mark.parametrize(
    "smart_config",
    [
        {"concurrencyPerSender": 0},
        {"concurrencyPerSender": 21},
        {"batchSize": 0},
        {"batchSize": 501},
        {"maxRetries": 11},
    ],
)
def test_out_of_range_smart_config_is_rejected(smart_config):
    with pytest.raises(ValidationError):
        TaskCreate.model_validate(_payload(smart_config))
