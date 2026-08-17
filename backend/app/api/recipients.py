import csv
import io
import zipfile
import pandas as pd
from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from pydantic import BaseModel, EmailStr, Field, TypeAdapter, ValidationError

from .deps import get_current_user
from ..config import settings
from ..models.user import User

router = APIRouter(prefix="/api/recipients", tags=["收件人"])
_EMAIL_ADAPTER = TypeAdapter(EmailStr)
_READ_CHUNK = 1024 * 1024


class RecipientItem(BaseModel):
    email: EmailStr
    name: str = Field(default="", max_length=200)


class ParseResponse(BaseModel):
    recipients: list[RecipientItem]
    total: int
    valid: int
    invalid: int


def is_valid_email(email: str) -> bool:
    try:
        _EMAIL_ADAPTER.validate_python(str(email).strip())
        return True
    except (ValidationError, TypeError, ValueError):
        return False


def _recipient(email: str, name: str = "") -> RecipientItem | None:
    try:
        return RecipientItem(email=email.strip(), name=name.strip()[:200])
    except ValidationError:
        return None


def parse_txt(content: str, max_items: int | None = None) -> list[RecipientItem]:
    recipients = []
    for line in content.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = line.split(",", 1)
        item = _recipient(parts[0], parts[1] if len(parts) > 1 else "")
        if item:
            recipients.append(item)
            if max_items and len(recipients) >= max_items:
                break
    return recipients


def parse_csv_content(content: str, max_items: int | None = None) -> list[RecipientItem]:
    recipients = []
    for row in csv.reader(io.StringIO(content)):
        if not row:
            continue
        item = _recipient(row[0], row[1] if len(row) > 1 else "")
        if item:
            recipients.append(item)
            if max_items and len(recipients) >= max_items:
                break
    return recipients


def parse_excel_content(content: bytes) -> list[RecipientItem]:
    recipients = []
    try:
        # XLSX is a ZIP container; reject decompression bombs before pandas
        # expands the workbook into memory. Legacy XLS files are checked by
        # the parser itself and remain subject to the upload-size limit.
        if content[:4] == b"PK\x03\x04":
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                infos = archive.infolist()
                if len(infos) > 2000 or sum(info.file_size for info in infos) > 100 * 1024 * 1024:
                    return []
        frame = pd.read_excel(io.BytesIO(content), header=None)
        for _, row in frame.iterrows():
            email = str(row.iloc[0]).strip() if len(row) else ""
            name = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
            item = _recipient(email, name)
            if item:
                recipients.append(item)
                if len(recipients) >= int(settings.MAX_RECIPIENTS_PER_IMPORT):
                    break
    except Exception:
        # Do not expose parser internals or stack traces to clients.
        return []
    return recipients


async def _read_limited(file: UploadFile, max_bytes: int) -> bytes:
    chunks: list[bytes] = []
    total = 0
    try:
        while chunk := await file.read(_READ_CHUNK):
            total += len(chunk)
            if total > max_bytes:
                raise HTTPException(status_code=413, detail="Recipient file is too large")
            chunks.append(chunk)
    finally:
        await file.close()
    return b"".join(chunks)


@router.post("/parse", response_model=ParseResponse)
async def parse_recipients(
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user),
):
    filename = (file.filename or "").lower()
    allowed = (".txt", ".csv", ".xlsx", ".xls")
    if not filename.endswith(allowed):
        raise HTTPException(status_code=400, detail="Unsupported file format")
    content_bytes = await _read_limited(
        file, max(1, int(settings.MAX_RECIPIENT_IMPORT_SIZE_MB)) * 1024 * 1024
    )

    if filename.endswith(".txt"):
        text = content_bytes.decode("utf-8", errors="ignore").lstrip("\ufeff")
        recipients = parse_txt(text, max_items=int(settings.MAX_RECIPIENTS_PER_IMPORT))
        total = sum(bool(line.strip()) for line in text.splitlines())
    elif filename.endswith(".csv"):
        text = content_bytes.decode("utf-8", errors="ignore").lstrip("\ufeff")
        recipients = parse_csv_content(text, max_items=int(settings.MAX_RECIPIENTS_PER_IMPORT))
        total = sum(1 for row in csv.reader(io.StringIO(text)) if row)
    else:
        recipients = parse_excel_content(content_bytes)
        total = len(recipients)

    if len(recipients) > int(settings.MAX_RECIPIENTS_PER_IMPORT):
        recipients = recipients[: int(settings.MAX_RECIPIENTS_PER_IMPORT)]
    valid = len(recipients)
    return ParseResponse(
        recipients=recipients,
        total=max(total, valid),
        valid=valid,
        invalid=max(0, total - valid),
    )


@router.post("/validate")
def validate_emails(emails: list[str], current_user: User = Depends(get_current_user)):
    if len(emails) > int(settings.MAX_RECIPIENTS_PER_IMPORT):
        raise HTTPException(status_code=413, detail="Too many email addresses")
    return [{"email": email, "valid": is_valid_email(email)} for email in emails]
