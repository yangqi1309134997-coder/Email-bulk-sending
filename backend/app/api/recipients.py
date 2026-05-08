import io
import csv
from typing import Optional
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File
from pydantic import BaseModel
from sqlmodel import Session
from .deps import get_current_user
from ..database import get_session
from ..models.user import User

router = APIRouter(prefix="/api/recipients", tags=["收件人"])


class RecipientItem(BaseModel):
    email: str
    name: str = ""


class ParseResponse(BaseModel):
    recipients: list[RecipientItem]
    total: int
    valid: int
    invalid: int


def is_valid_email(email: str) -> bool:
    return "@" in email and "." in email.split("@")[-1]


def parse_txt(content: str) -> list[RecipientItem]:
    recipients = []
    for line in content.strip().splitlines():
        line = line.strip()
        if not line or "@" not in line:
            continue
        parts = line.split(",", 1)
        email = parts[0].strip()
        name = parts[1].strip() if len(parts) > 1 else ""
        if is_valid_email(email):
            recipients.append(RecipientItem(email=email, name=name))
    return recipients


def parse_csv_content(content: str) -> list[RecipientItem]:
    recipients = []
    reader = csv.reader(io.StringIO(content))
    for row in reader:
        if not row:
            continue
        email = row[0].strip()
        name = row[1].strip() if len(row) > 1 else ""
        if is_valid_email(email):
            recipients.append(RecipientItem(email=email, name=name))
    return recipients


def parse_excel_content(content: bytes) -> list[RecipientItem]:
    import pandas as pd
    recipients = []
    try:
        df = pd.read_excel(io.BytesIO(content), header=None)
        for _, row in df.iterrows():
            email = str(row.iloc[0]).strip()
            name = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
            if is_valid_email(email):
                recipients.append(RecipientItem(email=email, name=name))
    except Exception:
        pass
    return recipients


@router.post("/parse", response_model=ParseResponse)
async def parse_recipients(file: UploadFile = File(...), current_user: User = Depends(get_current_user)):
    filename = file.filename or ""
    content_bytes = await file.read()

    recipients: list[RecipientItem] = []

    if filename.endswith(".txt"):
        text = content_bytes.decode("utf-8", errors="ignore")
        recipients = parse_txt(text)
    elif filename.endswith(".csv"):
        text = content_bytes.decode("utf-8", errors="ignore")
        recipients = parse_csv_content(text)
    elif filename.endswith((".xlsx", ".xls")):
        recipients = parse_excel_content(content_bytes)
    else:
        raise HTTPException(status_code=400, detail="Unsupported file format. Use .txt, .csv, .xlsx, or .xls")

    valid = len(recipients)
    return ParseResponse(recipients=recipients, total=valid, valid=valid, invalid=0)


@router.post("/validate")
def validate_emails(emails: list[str], current_user: User = Depends(get_current_user)):
    results = []
    for email in emails:
        results.append({"email": email, "valid": is_valid_email(email)})
    return results