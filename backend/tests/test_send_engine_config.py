"""Tests for send engine risk control helpers and smart_config parsing."""

from app.services.send_engine import _parse_smart_config
from app.services.email_sender import is_rate_limit_error, is_auth_error


def test_parse_smart_config_dict():
    assert _parse_smart_config({"a": 1})["a"] == 1


def test_parse_smart_config_json():
    assert _parse_smart_config('{"max_retries": 5}')["max_retries"] == 5


def test_parse_smart_config_invalid():
    assert _parse_smart_config("not-json") == {}
    assert _parse_smart_config(None) == {}
    assert _parse_smart_config("[1, 2, 3]") == {}


def test_risk_patterns():
    assert is_rate_limit_error("421 4.7.0 Try again later, rate limit")
    assert is_rate_limit_error("554 Message rejected as spam")
    assert is_auth_error("535 5.7.8 Username and Password not accepted")
