"""Tests for Aliyun DM signing helpers."""

from app.services.aliyun_dm import AliyunDMSender
from app.utils.security import encrypt_password
from types import SimpleNamespace


def test_percent_encode():
    s = AliyunDMSender()
    assert s._percent_encode("a b") == "a%20b"
    assert s._percent_encode("*") == "%2A"
    assert s._percent_encode("~") == "~"


def test_sign_deterministic():
    s = AliyunDMSender()
    params = {"Action": "SingleSendMail", "Format": "JSON", "Version": "2015-11-23"}
    sig1 = s._sign(params, "secret")
    sig2 = s._sign(params, "secret")
    assert sig1 == sig2
    assert isinstance(sig1, str) and len(sig1) > 10


def test_china_region_uses_canonical_directmail_endpoint(monkeypatch):
    sender = AliyunDMSender()
    captured = {}

    class Response:
        def __enter__(self):
            return self

        def __exit__(self, *args):
            return False

        def read(self):
            return b'{"RequestId":"req-1"}'

    def fake_urlopen(request, timeout):
        captured["url"] = request.full_url
        captured["timeout"] = timeout
        return Response()

    monkeypatch.setattr("urllib.request.urlopen", fake_urlopen)
    result = sender._call_api(
        "DescAccountSummary",
        {},
        SimpleNamespace(
            aliyun_access_key="key",
            aliyun_access_secret=encrypt_password("secret"),
            password="",
            aliyun_region="cn-hangzhou",
        ),
    )

    # Avoid coupling this test to the encryption key while still exercising endpoint selection.
    assert result["RequestId"] == "req-1"
    assert captured["url"].startswith("https://dm.aliyuncs.com/")


def test_tampered_encrypted_secret_fails_closed():
    sender = AliyunDMSender()

    result = sender._call_api(
        "DescAccountSummary",
        {},
        SimpleNamespace(
            id=7,
            aliyun_access_key="key",
            aliyun_access_secret="v3:invalid",
            password="",
            aliyun_region="cn-hangzhou",
        ),
    )

    assert result["Code"] == "MissingCredentials"


def test_directmail_retries_transient_throttling(monkeypatch):
    sender = AliyunDMSender()
    calls = []

    def fake_call(action, params, account):
        calls.append(action)
        if len(calls) == 1:
            return {"Code": "Throttling.User", "Message": "try again later"}
        return {"RequestId": "req-2"}

    monkeypatch.setattr(sender, "_call_api", fake_call)
    monkeypatch.setattr("app.services.aliyun_dm.time.sleep", lambda _: None)
    ok, message = sender.send(
        SimpleNamespace(email="from@example.com", aliyun_from_name=""),
        "to@example.com",
        "Recipient",
        "Subject",
        "<p>Body</p>",
    )

    assert ok is True
    assert message == "OK"
    assert calls == ["SingleSendMail", "SingleSendMail"]
