"""Tests for SMTP presets completeness."""

from app.utils.smtp_presets import (
    SMTP_PRESETS,
    get_preset,
    get_preset_for_email,
    get_preset_choices,
)


def test_major_providers_present():
    required = [
        "qq", "163", "126", "gmail", "outlook", "yahoo", "icloud",
        "aliyun_dm", "aliyun_enterprise", "tencent_exmail", "huawei_enterprise",
        "feishu", "dingtalk", "sendgrid", "mailgun", "ses", "postmark",
        "brevo", "mailjet", "custom",
    ]
    for key in required:
        assert key in SMTP_PRESETS, f"missing preset {key}"
        p = SMTP_PRESETS[key]
        assert p.name
        if key != "custom":
            assert p.smtp_server or key == "aliyun_dm"


def test_get_preset_for_email_qq():
    p = get_preset_for_email("user@qq.com")
    assert p is not None
    assert "qq" in p.smtp_server or p.name.startswith("QQ")


def test_preset_choices_serializable():
    choices = get_preset_choices()
    assert isinstance(choices, list) and len(choices) > 20
    assert all("key" in c and "name" in c for c in choices)


def test_get_preset_unknown():
    assert get_preset("no_such_provider") is None
