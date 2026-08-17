"""Alibaba Cloud DirectMail (阿里云邮箱推送) integration."""

from __future__ import annotations

import base64
import hashlib
import hmac
import html
import json
import logging
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
import random
import re
from typing import Tuple

logger = logging.getLogger(__name__)


class AliyunDMSender:
    """Send emails via Alibaba Cloud DirectMail API (SingleSendMail)."""

    API_VERSION = "2015-11-23"
    SIGNATURE_METHOD = "HMAC-SHA1"
    FORMAT = "JSON"

    @staticmethod
    def _endpoint_for_region(region: str) -> str:
        """Use Alibaba's canonical mainland endpoint; regional endpoints for overseas regions."""
        normalized = (region or "cn-hangzhou").strip().lower()
        # Region is user-configurable, but it must never become a hostname
        # injection vector. Unknown well-formed regions are still allowed for
        # newly launched Alibaba regions.
        if not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", normalized):
            normalized = "cn-hangzhou"
        mainland_regions = {
            "cn-hangzhou", "cn-shanghai", "cn-beijing", "cn-shenzhen",
            "cn-qingdao", "cn-zhangjiakou", "cn-huhehaote", "cn-chengdu",
        }
        if normalized in mainland_regions:
            return "dm.aliyuncs.com"
        return f"dm.{normalized}.aliyuncs.com"

    def _percent_encode(self, s: str) -> str:
        return (
            urllib.parse.quote(str(s), safe="~")
            .replace("+", "%20")
            .replace("*", "%2A")
        )

    def _sign(self, params: dict, access_secret: str) -> str:
        sorted_params = sorted(params.items())
        query_string = "&".join(
            f"{self._percent_encode(k)}={self._percent_encode(v)}" for k, v in sorted_params
        )
        string_to_sign = f"POST&{self._percent_encode('/')}&{self._percent_encode(query_string)}"
        key = (access_secret + "&").encode("utf-8")
        sign = hmac.new(key, string_to_sign.encode("utf-8"), hashlib.sha1).digest()
        return base64.b64encode(sign).decode("utf-8")

    def _credentials(self, sender) -> tuple[str, str, str]:
        from ..utils.security import decrypt_password

        access_key = (sender.aliyun_access_key or "").strip()
        secret_raw = sender.aliyun_access_secret or sender.password or ""
        access_secret: str | None = None
        if secret_raw:
            try:
                access_secret = decrypt_password(secret_raw)
            except Exception:
                # Plaintext is accepted only for legacy rows without an
                # encryption prefix. A tampered v2/v3 value must fail closed.
                if str(secret_raw).startswith(("v2:", "v3:")):
                    logger.warning("Invalid encrypted Aliyun secret for sender %s", getattr(sender, "id", "?"))
                    access_secret = None
                else:
                    access_secret = secret_raw
        region = (sender.aliyun_region or "cn-hangzhou").strip() or "cn-hangzhou"
        return access_key, access_secret or "", region

    def _call_api(self, action: str, params: dict, sender) -> dict:
        access_key, access_secret, region = self._credentials(sender)
        if not access_key or not access_secret:
            return {"Code": "MissingCredentials", "Message": "AccessKey 或 Secret 未配置"}

        endpoint = self._endpoint_for_region(region)
        public_params = {
            "Format": self.FORMAT,
            "Version": self.API_VERSION,
            "AccessKeyId": access_key,
            "SignatureMethod": self.SIGNATURE_METHOD,
            "Timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
            "SignatureVersion": "1.0",
            "SignatureNonce": str(uuid.uuid4()),
            "Action": action,
            "RegionId": region,
        }

        all_params = {**public_params, **params}
        all_params = {k: str(v) for k, v in all_params.items() if v is not None}
        signature = self._sign(all_params, access_secret)
        all_params["Signature"] = signature

        url = f"https://{endpoint}/"
        data = urllib.parse.urlencode(all_params).encode("utf-8")
        req = urllib.request.Request(url, data=data, method="POST")
        req.add_header("Content-Type", "application/x-www-form-urlencoded")

        try:
            # The URL is constructed from a validated region and a fixed HTTPS
            # scheme; it never accepts a caller-supplied URL or alternate scheme.
            with urllib.request.urlopen(req, timeout=30) as resp:  # nosec B310
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")
            logger.error("Aliyun DM API error: %s %s", e.code, body)
            try:
                return json.loads(body) if body else {"Code": str(e.code), "Message": body}
            except Exception:
                return {"Code": str(e.code), "Message": body}
        except Exception as e:
            logger.error("Aliyun DM API call failed: %s", e)
            return {"Code": "NetworkError", "Message": str(e)}

    def test_connection(self, sender) -> dict:
        """Test credentials. DescAccount is preferred; fall back to soft validation."""
        result = self._call_api("DescAccountSummary", {}, sender)
        if "Code" in result and result.get("Code") not in (None, "OK"):
            # Some accounts may not allow DescAccountSummary; try SenderStatisticsByTagNameAndBatchID soft-fail
            msg = result.get("Message") or result.get("Code")
            # MissingAction is still auth-valid in some cases; try a no-op style param validation via invalid action
            if result.get("Code") in ("InvalidAccessKeyId.NotFound", "SignatureDoesNotMatch", "IncompleteSignature", "MissingCredentials"):
                return {"success": False, "message": f"阿里云验证失败: {msg}"}
            # Other codes might mean API not enabled but key works
            if result.get("Code") in ("Forbidden.RAM", "NoPermission"):
                return {"success": True, "message": "AccessKey 有效（无 DescAccount 权限，可尝试发送）"}
            # Unknown action on older APIs
            if "InvalidAction" in str(result.get("Code", "")) or "not found" in str(msg).lower():
                return {"success": True, "message": "凭据格式有效（接口权限受限）"}
            return {"success": False, "message": f"阿里云验证失败: {msg}"}
        return {"success": True, "message": "阿里云邮箱推送连接成功"}

    def send(
        self,
        sender,
        recipient_email: str,
        recipient_name: str,
        subject: str,
        body_html: str,
        from_name: str = "",
        max_retries: int | None = None,
        retry_backoff_base: float | None = None,
    ) -> Tuple[bool, str]:
        name = recipient_name or ""
        from_alias = from_name or getattr(sender, "aliyun_from_name", "") or ""
        to_address = recipient_email
        # DirectMail ToAddress does not always accept display-name form; keep plain email for reliability
        params = {
            "AccountName": sender.email,
            "ReplyToAddress": "true",
            "AddressType": "1",
            "ToAddress": to_address,
            "FromAlias": from_alias,
            "Subject": subject.replace("{name}", name or "朋友").replace("{email}", recipient_email),
            "HtmlBody": (body_html or "")
            .replace("{name}", html.escape(name or "朋友"))
            .replace("{email}", html.escape(recipient_email, quote=True)),
        }

        from ..config import settings

        if max_retries is None:
            max_retries = int(getattr(settings, "ALIYUN_MAX_RETRIES", 3))
        if retry_backoff_base is None:
            retry_backoff_base = float(
                getattr(settings, "ALIYUN_RETRY_BACKOFF_BASE", 1.0)
            )
        max_retries = max(0, min(10, int(max_retries)))
        backoff_base = max(0.1, min(60.0, float(retry_backoff_base)))
        result = {"Code": "NetworkError", "Message": "not attempted"}
        for attempt in range(max_retries + 1):
            result = self._call_api("SingleSendMail", params, sender)
            if not self._is_transient(result) or attempt >= max_retries:
                break
            # Retry jitter is intentionally non-cryptographic.
            delay = min(
                3600.0,
                backoff_base * (2 ** attempt) + random.uniform(0, backoff_base),  # nosec B311
            )
            time.sleep(delay)
        if "Code" in result and result.get("Code") not in (None, "OK"):
            code = result.get("Code", "unknown")
            message = result.get("Message", "")
            return False, f"Aliyun DM error: {code} - {message}"
        # Success responses typically contain RequestId / EnvId without Code
        return True, "OK"

    @staticmethod
    def _is_transient(result: dict) -> bool:
        code = str((result or {}).get("Code", "")).lower()
        message = str((result or {}).get("Message", "")).lower()
        transient_codes = (
            "networkerror", "throttl", "too many", "internal", "serviceunavailable",
            "requesttimeout", "timeout", "temporarily", "429", "500", "502", "503", "504",
        )
        return any(token in code or token in message for token in transient_codes)


aliyun_dm_sender = AliyunDMSender()
