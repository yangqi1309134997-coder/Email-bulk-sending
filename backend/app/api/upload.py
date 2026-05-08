import os
import uuid
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File
from .deps import get_current_user
from ..models.user import User
from ..config import settings

router = APIRouter(prefix="/api/upload", tags=["文件上传"])

ALLOWED_ATTACHMENTS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
    ".txt", ".csv", ".zip", ".rar", ".png", ".jpg", ".jpeg", ".gif", ".bmp",
}

ALLOWED_IMAGES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg", ".webp"}


@router.post("/attachment")
async def upload_attachment(file: UploadFile = File(...), current_user: User = Depends(get_current_user)):
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in ALLOWED_ATTACHMENTS:
        raise HTTPException(status_code=400, detail=f"File type {ext} not allowed")

    content = await file.read()
    if len(content) > settings.MAX_ATTACHMENT_SIZE_MB * 1024 * 1024:
        raise HTTPException(status_code=400, detail=f"File too large (max {settings.MAX_ATTACHMENT_SIZE_MB}MB)")

    file_id = uuid.uuid4().hex[:8]
    safe_name = f"{file_id}_{file.filename}"
    upload_dir = os.path.join(settings.UPLOAD_DIR, "attachments")
    os.makedirs(upload_dir, exist_ok=True)
    file_path = os.path.join(upload_dir, safe_name)

    with open(file_path, "wb") as f:
        f.write(content)

    return {"filename": file.filename, "path": file_path, "size": len(content)}


@router.post("/image")
async def upload_image(file: UploadFile = File(...), current_user: User = Depends(get_current_user)):
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in ALLOWED_IMAGES:
        raise HTTPException(status_code=400, detail=f"Image type {ext} not allowed")

    content = await file.read()
    if len(content) > settings.MAX_IMAGE_SIZE_MB * 1024 * 1024:
        raise HTTPException(status_code=400, detail=f"Image too large (max {settings.MAX_IMAGE_SIZE_MB}MB)")

    file_id = uuid.uuid4().hex[:8]
    safe_name = f"{file_id}_{file.filename}"
    upload_dir = os.path.join(settings.UPLOAD_DIR, "images")
    os.makedirs(upload_dir, exist_ok=True)
    file_path = os.path.join(upload_dir, safe_name)

    with open(file_path, "wb") as f:
        f.write(content)

    return {"filename": file.filename, "path": file_path, "size": len(content)}