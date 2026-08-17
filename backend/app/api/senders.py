import json
from datetime import datetime
from typing import Literal, Optional
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel, ConfigDict, EmailStr, Field, field_validator, model_validator
from sqlalchemy.exc import IntegrityError
from sqlmodel import Session, select
from .deps import get_current_user
from ..database import get_session
from ..models.send_log import SendLog
from ..models.sender import Sender
from ..models.task import Task
from ..models.user import User
from ..utils.security import encrypt_password
from ..utils.smtp_presets import get_preset_choices, get_preset
from ..utils.time import utcnow

router = APIRouter(prefix="/api/senders", tags=["发件人"])


def _task_references_sender(raw_sender_ids: str, sender_id: int) -> bool:
    try:
        values = json.loads(raw_sender_ids or "[]")
    except (TypeError, json.JSONDecodeError):
        return False
    if not isinstance(values, list):
        return False
    for value in values:
        if isinstance(value, bool):
            continue
        if isinstance(value, int) and value == sender_id:
            return True
        if isinstance(value, str) and value.strip().isascii() and value.strip().isdecimal():
            try:
                if int(value.strip()) == sender_id:
                    return True
            except ValueError:
                continue
    return False


def _validated_smtp_server(value: str) -> str:
    if not value:
        return value
    from ..services.email_sender import _normalize_smtp_host

    try:
        return _normalize_smtp_host(value)
    except ValueError:
        raise ValueError("SMTP服务器地址不合法") from None


def _validated_display_text(value):
    if value is not None and any(ord(char) < 32 or ord(char) == 127 for char in value):
        raise ValueError("显示名称不能包含控制字符")
    return value


class SenderCreate(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    email: EmailStr
    password: str = Field(default="", max_length=512)
    smtp_server: str = Field(default="", max_length=255)
    smtp_port: int = Field(default=587, ge=0, le=65535)
    use_tls: bool = True
    smtp_username: str = Field(default="", max_length=320)
    smtp_security: Literal["", "ssl", "starttls", "none"] = ""
    sender_type: str = Field(default="自定义SMTP", max_length=64)
    enabled: bool = True
    weight: int = Field(default=50, ge=1, le=100)
    daily_quota: int = Field(default=500, ge=0, le=10_000_000)
    # 阿里云邮箱推送
    aliyun_access_key: str = Field(default="", max_length=256)
    aliyun_access_secret: str = Field(default="", max_length=512)
    aliyun_region: str = Field(default="cn-hangzhou", max_length=64, pattern=r"^[A-Za-z0-9-]+$")
    aliyun_from_name: str = Field(default="", max_length=128)

    _normalize_smtp_server = field_validator("smtp_server")(_validated_smtp_server)
    _validate_aliyun_from_name = field_validator("aliyun_from_name")(_validated_display_text)


class SenderUpdate(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    email: Optional[EmailStr] = None
    password: Optional[str] = Field(default=None, max_length=512)
    smtp_server: Optional[str] = Field(default=None, max_length=255)
    smtp_port: Optional[int] = Field(default=None, ge=0, le=65535)
    use_tls: Optional[bool] = None
    smtp_username: Optional[str] = Field(default=None, max_length=320)
    smtp_security: Optional[Literal["", "ssl", "starttls", "none"]] = None
    sender_type: Optional[str] = Field(default=None, max_length=64)
    enabled: Optional[bool] = None
    weight: Optional[int] = Field(default=None, ge=1, le=100)
    daily_quota: Optional[int] = Field(default=None, ge=0, le=10_000_000)
    aliyun_access_key: Optional[str] = Field(default=None, max_length=256)
    aliyun_access_secret: Optional[str] = Field(default=None, max_length=512)
    aliyun_region: Optional[str] = Field(default=None, max_length=64, pattern=r"^[A-Za-z0-9-]+$")
    aliyun_from_name: Optional[str] = Field(default=None, max_length=128)
    status: Optional[Literal["active", "paused", "banned"]] = None

    _normalize_smtp_server = field_validator("smtp_server")(_validated_smtp_server)
    _validate_aliyun_from_name = field_validator("aliyun_from_name")(_validated_display_text)


class SenderResponse(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    email: str
    smtp_server: str
    smtp_port: int
    use_tls: bool
    smtp_username: str = ""
    smtp_security: str = ""
    sender_type: str
    enabled: bool
    weight: int
    daily_quota: int
    daily_sent: int
    success_rate: float
    status: str
    aliyun_access_key: str = ""
    aliyun_region: str = "cn-hangzhou"
    aliyun_from_name: str = ""
    created_at: datetime
    updated_at: datetime

# Sender type constants
SENDER_TYPES = [
    "自定义SMTP",
    "QQ邮箱",
    "163邮箱",
    "126邮箱",
    "yeah邮箱",
    "新浪邮箱",
    "搜狐邮箱",
    "Naver邮箱",
    "Daum / Hanmail",
    "139邮箱",
    "189邮箱",
    "Gmail",
    "Outlook",
    "Yahoo",
    "iCloud",
    "Zoho",
    "AOL",
    "Fastmail",
    "Microsoft 365",
    "Mail.com",
    "T-Online",
    "Orange Mail",
    "Bluewin",
    "Yandex",
    "阿里云邮箱推送",
    "aliyun_dm",
    "阿里云邮箱推送SMTP",
    "aliyun_dm_smtp",
    "阿里企业邮箱",
    "腾讯企业邮箱",
    "华为企业邮箱",
    "网易企业邮箱",
    "飞书邮箱",
    "Lark Mail",
    "钉钉邮箱",
    "SendGrid",
    "Mailgun",
    "Amazon SES",
    "Postmark",
    "SparkPost",
    "Brevo",
    "Elastic Email",
    "Mailjet",
]


# Map UI sender_type labels to built-in preset keys
SENDER_TYPE_PRESET_MAP = {
    "QQ邮箱": "qq_ssl",
    "163邮箱": "163",
    "126邮箱": "126",
    "yeah邮箱": "yeah",
    "新浪邮箱": "sina",
    "搜狐邮箱": "sohu",
    "Naver邮箱": "naver",
    "Daum / Hanmail": "daum",
    "139邮箱": "139",
    "189邮箱": "189",
    "Gmail": "gmail",
    "Outlook": "outlook",
    "Yahoo": "yahoo",
    "iCloud": "icloud",
    "Zoho": "zoho",
    "AOL": "aol",
    "Fastmail": "fastmail",
    "Microsoft 365": "office365",
    "Mail.com": "mail_com",
    "T-Online": "t_online",
    "Orange Mail": "orange",
    "Bluewin": "bluewin",
    "Yandex": "yandex",
    "阿里企业邮箱": "aliyun_enterprise",
    "腾讯企业邮箱": "tencent_exmail",
    "华为企业邮箱": "huawei_enterprise",
    "网易企业邮箱": "netease_enterprise",
    "飞书邮箱": "feishu",
    "Lark Mail": "larksuite",
    "钉钉邮箱": "dingtalk",
    "阿里云邮箱推送SMTP": "aliyun_dm_587",
    "SendGrid": "sendgrid",
    "Mailgun": "mailgun",
    "Amazon SES": "ses",
    "Postmark": "postmark",
    "SparkPost": "sparkpost",
    "Brevo": "brevo",
    "Elastic Email": "elastic_email",
    "Mailjet": "mailjet",
}

# Keep the API whitelist and preset catalog in sync. Explicit mappings above
# select preferred variants (for example QQ STARTTLS); every remaining catalog
# entry is accepted automatically so the UI cannot offer an unsavable choice.
for _choice in get_preset_choices():
    _sender_type = _choice.get("sender_type")
    _preset_key = _choice.get("key")
    if _sender_type and _sender_type not in SENDER_TYPES:
        SENDER_TYPES.append(_sender_type)
    if _sender_type and _preset_key:
        SENDER_TYPE_PRESET_MAP.setdefault(_sender_type, _preset_key)


@router.get("", response_model=list[SenderResponse])
def list_senders(current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    senders = session.exec(select(Sender).where(Sender.user_id == current_user.id)).all()
    return senders


@router.get("/types", response_model=list[str])
def list_sender_types(current_user: User = Depends(get_current_user)):
    return SENDER_TYPES


@router.post("", response_model=SenderResponse)
def create_sender(req: SenderCreate, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    if req.sender_type not in SENDER_TYPES:
        raise HTTPException(status_code=400, detail="不支持的发件人类型")
    normalized_email = str(req.email).strip().lower()
    duplicate = session.exec(
        select(Sender).where(
            Sender.user_id == current_user.id,
            Sender.email == normalized_email,
        )
    ).first()
    if duplicate:
        raise HTTPException(status_code=400, detail="该发件人邮箱已存在")
    smtp_server = req.smtp_server
    smtp_port = req.smtp_port
    use_tls = req.use_tls
    smtp_username = req.smtp_username.strip()
    smtp_security = req.smtp_security
    password = req.password
    daily_quota = req.daily_quota

    # Auto-fill from preset map when server not provided
    preset_key = SENDER_TYPE_PRESET_MAP.get(req.sender_type)
    if preset_key and (not smtp_server or req.sender_type != "自定义SMTP"):
        preset = get_preset(preset_key)
        if preset:
            if not smtp_server:
                smtp_server = preset.smtp_server
            if not req.smtp_port or req.smtp_port == 587:
                smtp_port = preset.smtp_port
            use_tls = preset.use_tls
            if not smtp_security:
                smtp_security = getattr(preset, "security", "")
            if "daily_quota" not in req.model_fields_set:
                daily_quota = preset.daily_limit

    if req.sender_type in ("阿里云邮箱推送", "aliyun_dm", "Aliyun DM"):
        smtp_server = "dm.aliyuncs.com"
        smtp_port = 0
        smtp_security = "none"
        password = req.aliyun_access_secret or req.password or ""
    elif req.sender_type in ("阿里云邮箱推送SMTP", "aliyun_dm_smtp"):
        smtp_server = smtp_server or "smtpdm.aliyun.com"
        smtp_port = smtp_port or 465
        use_tls = True

    if not req.email:
        raise HTTPException(status_code=400, detail="邮箱地址不能为空")
    if req.sender_type not in ("阿里云邮箱推送", "aliyun_dm") and not password:
        raise HTTPException(status_code=400, detail="密码/授权码不能为空")
    if req.sender_type in ("阿里云邮箱推送", "aliyun_dm") and not req.aliyun_access_key:
        raise HTTPException(status_code=400, detail="AccessKey ID 不能为空")
    if req.sender_type in ("阿里云邮箱推送", "aliyun_dm") and not password:
        raise HTTPException(status_code=400, detail="AccessKey Secret 不能为空")
    if req.sender_type not in ("阿里云邮箱推送", "aliyun_dm") and smtp_port <= 0:
        raise HTTPException(status_code=400, detail="SMTP端口必须大于0")

    sender = Sender(
        user_id=current_user.id,
        email=normalized_email,
        password=encrypt_password(password or ""),
        smtp_server=smtp_server or "",
        smtp_port=smtp_port or 0,
        use_tls=use_tls,
        smtp_username=smtp_username,
        smtp_security=smtp_security,
        sender_type=req.sender_type,
        enabled=req.enabled,
        weight=req.weight,
        daily_quota=daily_quota,
        aliyun_access_key=req.aliyun_access_key,
        aliyun_access_secret=encrypt_password(req.aliyun_access_secret) if req.aliyun_access_secret else "",
        aliyun_region=req.aliyun_region or "cn-hangzhou",
        aliyun_from_name=req.aliyun_from_name,
    )
    session.add(sender)
    session.commit()
    session.refresh(sender)
    return sender


@router.put("/{sender_id}", response_model=SenderResponse)
def update_sender(sender_id: int, req: SenderUpdate, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    sender = session.get(Sender, sender_id)
    if not sender or sender.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Sender not found")

    update_data = req.model_dump(exclude_unset=True)
    if "email" in update_data:
        normalized_email = str(update_data["email"]).strip().lower()
        duplicate = session.exec(
            select(Sender).where(
                Sender.user_id == current_user.id,
                Sender.email == normalized_email,
                Sender.id != sender_id,
            )
        ).first()
        if duplicate:
            raise HTTPException(status_code=400, detail="该发件人邮箱已存在")
        update_data["email"] = normalized_email
    if update_data.get("sender_type") and update_data["sender_type"] not in SENDER_TYPES:
        raise HTTPException(status_code=400, detail="不支持的发件人类型")
    if "password" in update_data and update_data["password"]:
        update_data["password"] = encrypt_password(update_data.pop("password"))
    elif "password" in update_data:
        update_data.pop("password")
    if "aliyun_access_secret" in update_data and update_data["aliyun_access_secret"]:
        update_data["aliyun_access_secret"] = encrypt_password(update_data["aliyun_access_secret"])
    elif "aliyun_access_secret" in update_data:
        update_data.pop("aliyun_access_secret")

    if "sender_type" in update_data:
        new_type = update_data["sender_type"]
        if new_type in ("阿里云邮箱推送", "aliyun_dm"):
            update_data["smtp_server"] = "dm.aliyuncs.com"
            update_data["smtp_port"] = 0
        elif new_type in ("阿里云邮箱推送SMTP", "aliyun_dm_smtp"):
            update_data.setdefault("smtp_server", "smtpdm.aliyun.com")
            update_data.setdefault("smtp_port", 465)
        else:
            preset_key = SENDER_TYPE_PRESET_MAP.get(new_type)
            if preset_key:
                preset = get_preset(preset_key)
                if preset:
                    update_data.setdefault("smtp_server", preset.smtp_server)
                    update_data.setdefault("smtp_port", preset.smtp_port)
                    update_data.setdefault("use_tls", preset.use_tls)
                    update_data.setdefault("smtp_security", getattr(preset, "security", ""))

    resulting_type = update_data.get("sender_type", sender.sender_type)
    resulting_port = update_data.get("smtp_port", sender.smtp_port)
    resulting_server = update_data.get("smtp_server", sender.smtp_server)
    if resulting_type not in ("阿里云邮箱推送", "aliyun_dm"):
        if not resulting_server or int(resulting_port or 0) <= 0:
            raise HTTPException(status_code=400, detail="SMTP服务器和端口不能为空")

    for key, value in update_data.items():
        setattr(sender, key, value)

    sender.updated_at = utcnow()
    session.add(sender)
    session.commit()
    session.refresh(sender)
    return sender


@router.delete("/{sender_id}")
def delete_sender(sender_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    sender = session.get(Sender, sender_id)
    if not sender or sender.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Sender not found")

    active_task_sender_ids = session.exec(
        select(Task.sender_ids).where(
            Task.user_id == current_user.id,
            Task.status.in_(["pending", "running", "paused"]),  # type: ignore[attr-defined]
        )
    ).all()
    if any(_task_references_sender(raw, sender_id) for raw in active_task_sender_ids):
        raise HTTPException(
            status_code=409,
            detail="发件人正在被未完成任务使用，请先结束任务或禁用发件人",
        )

    has_history = session.exec(
        select(SendLog.id).where(SendLog.sender_id == sender_id).limit(1)
    ).first()
    if has_history is not None:
        raise HTTPException(
            status_code=409,
            detail="发件人已有发送历史，为保留审计记录请改为禁用",
        )

    session.delete(sender)
    try:
        session.commit()
    except IntegrityError:
        session.rollback()
        raise HTTPException(
            status_code=409,
            detail="发件人已被其他数据引用，请改为禁用",
        ) from None
    return {"message": "Sender deleted"}


@router.post("/{sender_id}/test")
def test_sender(sender_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    sender = session.get(Sender, sender_id)
    if not sender or sender.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Sender not found")

    # 阿里云邮箱推送使用 API 测试
    if sender.sender_type in ("阿里云邮箱推送", "aliyun_dm", "Aliyun DM"):
        try:
            from ..services.aliyun_dm import aliyun_dm_sender
            return aliyun_dm_sender.test_connection(sender)
        except Exception as e:
            return {"success": False, "message": f"阿里云邮箱推送测试失败: {str(e)}"}

    # 普通 SMTP / 阿里云 SMTP 使用连接池（正确处理 465 SSL）
    try:
        from ..services.email_sender import email_sender
        ok, message = email_sender.test_connection(sender)
        return {"success": ok, "message": message}
    except Exception as e:
        return {"success": False, "message": f"连接失败: {str(e)}"}


@router.post("/{sender_id}/toggle", response_model=SenderResponse)
def toggle_sender(sender_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    sender = session.get(Sender, sender_id)
    if not sender or sender.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Sender not found")
    sender.enabled = not sender.enabled
    sender.updated_at = utcnow()
    session.add(sender)
    session.commit()
    session.refresh(sender)
    return sender


@router.get("/presets", response_model=list[dict])
def get_smtp_presets(current_user: User = Depends(get_current_user)):
    """Get all SMTP presets for frontend dropdown."""
    return get_preset_choices()


@router.get("/presets/{preset_key}", response_model=dict)
def get_smtp_preset(preset_key: str, current_user: User = Depends(get_current_user)):
    """Get a specific SMTP preset by key."""
    preset = get_preset(preset_key)
    if not preset:
        raise HTTPException(status_code=404, detail="Preset not found")
    return {
        "key": preset_key,
        "name": preset.name,
        "type": preset.type,
        "server": preset.smtp_server,
        "port": preset.smtp_port,
        "tls": preset.use_tls,
        "security": preset.security,
        "daily_limit": preset.daily_limit,
        "batch_limit": preset.batch_limit,
        "delay_range": list(preset.delay_range),
        "notes": preset.notes,
        "domains": preset.domains,
        "requires_app_password": preset.requires_app_password,
        "auth_type": preset.auth_type,
        "requires_local_service": preset.requires_local_service,
    }


# Configuration Templates / History
class SenderTemplateCreate(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    name: str = Field(min_length=1, max_length=200)
    description: str = Field(default="", max_length=1000)
    sender_type: str = Field(min_length=1, max_length=64)
    smtp_server: str = Field(default="", max_length=255)
    smtp_port: int = Field(default=587, ge=0, le=65535)
    use_tls: bool = True
    smtp_username: str = Field(default="", max_length=320)
    smtp_security: Literal["", "ssl", "starttls", "none"] = ""
    weight: int = Field(default=50, ge=1, le=100)
    daily_quota: int = Field(default=500, ge=0, le=10_000_000)
    aliyun_access_key: str = Field(default="", max_length=256)
    aliyun_access_secret: str = Field(default="", max_length=512)
    aliyun_region: str = Field(default="cn-hangzhou", max_length=64, pattern=r"^[A-Za-z0-9-]+$")
    aliyun_from_name: str = Field(default="", max_length=128)

    _normalize_smtp_server = field_validator("smtp_server")(_validated_smtp_server)
    _validate_aliyun_from_name = field_validator("aliyun_from_name")(_validated_display_text)

    @field_validator("sender_type")
    @classmethod
    def validate_sender_type(cls, value: str) -> str:
        if value not in SENDER_TYPES:
            raise ValueError("不支持的发件人类型")
        return value

    @model_validator(mode="after")
    def validate_transport(self):
        if self.sender_type not in ("阿里云邮箱推送", "aliyun_dm"):
            if not self.smtp_server or self.smtp_port <= 0:
                raise ValueError("SMTP服务器和端口不能为空")
        return self


class SenderTemplateResponse(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    name: str
    description: str
    sender_type: str
    smtp_server: str
    smtp_port: int
    use_tls: bool
    smtp_username: str = ""
    smtp_security: str = ""
    weight: int
    daily_quota: int
    aliyun_access_key: str
    aliyun_region: str
    aliyun_from_name: str
    created_at: datetime
    updated_at: datetime

@router.post("/templates", response_model=SenderTemplateResponse)
def create_sender_template(req: SenderTemplateCreate, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    """Save a sender configuration as a template for reuse."""
    from ..models.sender_template import SenderTemplate
    template = SenderTemplate(
        user_id=current_user.id,
        name=req.name.strip(),
        description=req.description,
        sender_type=req.sender_type,
        smtp_server=req.smtp_server,
        smtp_port=req.smtp_port,
        use_tls=req.use_tls,
        smtp_username=req.smtp_username,
        smtp_security=req.smtp_security,
        weight=req.weight,
        daily_quota=req.daily_quota,
        aliyun_access_key=req.aliyun_access_key,
        aliyun_access_secret=encrypt_password(req.aliyun_access_secret) if req.aliyun_access_secret else "",
        aliyun_region=req.aliyun_region,
        aliyun_from_name=req.aliyun_from_name,
    )
    session.add(template)
    session.commit()
    session.refresh(template)
    return template


@router.get("/templates", response_model=list[SenderTemplateResponse])
def list_sender_templates(current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    """List all saved sender templates."""
    from ..models.sender_template import SenderTemplate
    templates = session.exec(
        select(SenderTemplate)
        .where(SenderTemplate.user_id == current_user.id)
        .order_by(SenderTemplate.updated_at.desc(), SenderTemplate.id.desc())
        .limit(500)
    ).all()
    return templates


@router.delete("/templates/{template_id}")
def delete_sender_template(template_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    """Delete a sender template."""
    from ..models.sender_template import SenderTemplate
    template = session.get(SenderTemplate, template_id)
    if not template or template.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Template not found")
    session.delete(template)
    session.commit()
    return {"message": "Template deleted"}


class ApplyTemplateRequest(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    email: EmailStr
    password: str = Field(default="", max_length=512)
    aliyun_access_key: str = Field(default="", max_length=256)
    aliyun_access_secret: str = Field(default="", max_length=512)


@router.post("/templates/{template_id}/apply", response_model=SenderResponse)
def apply_sender_template(
    template_id: int,
    req: ApplyTemplateRequest,
    current_user: User = Depends(get_current_user),
    session: Session = Depends(get_session),
):
    """Apply a saved template to create a new sender."""
    from ..models.sender_template import SenderTemplate
    template = session.get(SenderTemplate, template_id)
    if not template or template.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Template not found")

    normalized_email = str(req.email).strip().lower()
    duplicate = session.exec(
        select(Sender).where(
            Sender.user_id == current_user.id,
            Sender.email == normalized_email,
        )
    ).first()
    if duplicate:
        raise HTTPException(status_code=400, detail="该发件人邮箱已存在")

    smtp_server = template.smtp_server
    smtp_port = template.smtp_port
    use_tls = template.use_tls
    password_enc = encrypt_password(req.password) if req.password else encrypt_password("")
    aliyun_key = req.aliyun_access_key or template.aliyun_access_key
    aliyun_secret = (req.aliyun_access_secret or "").strip()
    # Frontend may accidentally resubmit already-encrypted secrets from list API.
    if aliyun_secret.startswith(("v2:", "v3:")):
        aliyun_secret_enc = aliyun_secret
    elif aliyun_secret:
        aliyun_secret_enc = encrypt_password(aliyun_secret)
    else:
        aliyun_secret_enc = template.aliyun_access_secret or ""

    if template.sender_type in ("阿里云邮箱推送", "aliyun_dm"):
        smtp_server = "dm.aliyuncs.com"
        smtp_port = 0
        # password field mirrors access secret ciphertext
        password_enc = aliyun_secret_enc or password_enc
        if not aliyun_key or not aliyun_secret_enc:
            raise HTTPException(status_code=400, detail="阿里云邮箱推送模板缺少有效凭据")
    elif template.sender_type in ("阿里云邮箱推送SMTP", "aliyun_dm_smtp"):
        smtp_server = smtp_server or "smtpdm.aliyun.com"
        smtp_port = smtp_port or 465
        use_tls = True

    if template.sender_type not in ("阿里云邮箱推送", "aliyun_dm"):
        if not req.password:
            raise HTTPException(status_code=400, detail="密码/授权码不能为空")
        if not smtp_server or smtp_port <= 0:
            raise HTTPException(status_code=400, detail="SMTP服务器和端口不能为空")

    sender = Sender(
        user_id=current_user.id,
        email=normalized_email,
        password=password_enc,
        smtp_server=smtp_server,
        smtp_port=smtp_port,
        use_tls=use_tls,
        smtp_username=template.smtp_username,
        smtp_security=template.smtp_security,
        sender_type=template.sender_type,
        enabled=True,
        weight=template.weight,
        daily_quota=template.daily_quota,
        aliyun_access_key=aliyun_key,
        aliyun_access_secret=aliyun_secret_enc,
        aliyun_region=template.aliyun_region,
        aliyun_from_name=template.aliyun_from_name,
    )
    session.add(sender)
    session.commit()
    session.refresh(sender)
    return sender


@router.post("/from-preset/{preset_key}", response_model=SenderResponse)
def create_sender_from_preset(
    preset_key: str,
    req: SenderCreate,
    current_user: User = Depends(get_current_user),
    session: Session = Depends(get_session),
):
    """Create sender using a built-in SMTP preset."""
    preset = get_preset(preset_key)
    if not preset:
        raise HTTPException(status_code=404, detail="Preset not found")

    explicit_sender_type = (
        req.sender_type if "sender_type" in req.model_fields_set else None
    )
    preset_choice = next(
        (choice for choice in get_preset_choices() if choice.get("key") == preset_key),
        None,
    )
    sender_type = (
        explicit_sender_type
        or (preset_choice or {}).get("sender_type")
        or (preset.name if preset.type != "custom" else "自定义SMTP")
    )
    smtp_server = preset.smtp_server or req.smtp_server
    smtp_port = preset.smtp_port or req.smtp_port
    use_tls = preset.use_tls if preset.smtp_server else req.use_tls
    password = req.password

    if preset_key == "aliyun_dm" or sender_type in ("阿里云邮箱推送", "aliyun_dm"):
        smtp_server = "dm.aliyuncs.com"
        smtp_port = 0
        use_tls = False
        password = req.aliyun_access_secret or req.password
        sender_type = "阿里云邮箱推送"
        if not req.aliyun_access_key or not (req.aliyun_access_secret or req.password):
            raise HTTPException(status_code=400, detail="AccessKey ID/Secret 不能为空")
    elif smtp_port <= 0 or not smtp_server:
        raise HTTPException(status_code=400, detail="SMTP服务器和端口不能为空")
    elif not password:
        raise HTTPException(status_code=400, detail="密码/授权码不能为空")

    normalized_email = str(req.email).strip().lower()
    if session.exec(
        select(Sender).where(
            Sender.user_id == current_user.id,
            Sender.email == normalized_email,
        )
    ).first():
        raise HTTPException(status_code=400, detail="该发件人邮箱已存在")

    sender = Sender(
        user_id=current_user.id,
        email=normalized_email,
        password=encrypt_password(password or ""),
        smtp_server=smtp_server,
        smtp_port=smtp_port,
        use_tls=use_tls,
        smtp_username=req.smtp_username,
        smtp_security=("none" if sender_type == "阿里云邮箱推送" else getattr(preset, "security", "")),
        sender_type=sender_type,
        enabled=req.enabled,
        weight=req.weight if req.weight is not None else 50,
        daily_quota=(
            preset.daily_limit
            if "daily_quota" not in req.model_fields_set
            else req.daily_quota
        ),
        aliyun_access_key=req.aliyun_access_key,
        aliyun_access_secret=encrypt_password(req.aliyun_access_secret) if req.aliyun_access_secret else "",
        aliyun_region=req.aliyun_region,
        aliyun_from_name=req.aliyun_from_name,
    )
    session.add(sender)
    session.commit()
    session.refresh(sender)
    return sender
