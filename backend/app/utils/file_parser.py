import re
import io
import csv
from typing import Optional
from pydantic import BaseModel


class RecipientItem(BaseModel):
    email: str
    name: str = ""


class ParseResult(BaseModel):
    recipients: list[RecipientItem]
    total: int
    valid: int
    invalid: int


EMAIL_REGEX = re.compile(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")


def is_valid_email(email: str) -> bool:
    return bool(EMAIL_REGEX.match(email))


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


def parse_file(filename: str, content_bytes: bytes) -> ParseResult:
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
        raise ValueError("Unsupported file format. Use .txt, .csv, .xlsx, or .xls")

    valid = len(recipients)
    return ParseResult(recipients=recipients, total=valid, valid=valid, invalid=0)
