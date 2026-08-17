import os
import re
import uuid
from pathlib import Path

import aiofiles
from fastapi import APIRouter, Depends, File, HTTPException, UploadFile

from .deps import get_current_user
from ..config import settings
from ..models.user import User

router = APIRouter(prefix="/api/upload", tags=["文件上传"])

ALLOWED_ATTACHMENTS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
    ".txt", ".csv", ".zip", ".rar", ".png", ".jpg", ".jpeg", ".gif", ".bmp",
}
ALLOWED_IMAGES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
_CHUNK_SIZE = 1024 * 1024


def _safe_filename(filename: str) -> str:
    basename = Path((filename or "").replace("\\", "/")).name
    basename = re.sub(r"[\x00-\x1f\x7f]", "", basename).strip().strip(".")
    if not basename:
        raise HTTPException(status_code=400, detail="Invalid filename")
    stem = Path(basename).stem[:120].rstrip(" .") or "file"
    suffix = Path(basename).suffix.lower()[:16]
    return f"{stem}{suffix}"


def user_upload_root(user_id: int, category: str) -> Path:
    return (Path(settings.UPLOAD_DIR) / str(int(user_id)) / category).resolve()


def resolve_user_attachment_paths(user_id: int, paths: list[str]) -> list[str]:
    """Resolve existing attachment paths and enforce per-user ownership."""
    root = user_upload_root(user_id, "attachments")
    resolved: list[str] = []
    for supplied in paths:
        try:
            candidate = Path(str(supplied)).expanduser()
            if not candidate.is_absolute():
                candidate = root / candidate
            candidate = candidate.resolve(strict=True)
        except (OSError, RuntimeError, ValueError):
            raise ValueError("Attachment does not exist") from None
        if not candidate.is_file() or not candidate.is_relative_to(root):
            raise ValueError("Attachment is outside your upload directory")
        resolved.append(str(candidate))
    return resolved


def _looks_like_image(path: Path, extension: str) -> bool:
    with path.open("rb") as stream:
        header = stream.read(16)
    if extension == ".png":
        return header.startswith(b"\x89PNG\r\n\x1a\n")
    if extension in {".jpg", ".jpeg"}:
        return header.startswith(b"\xff\xd8\xff")
    if extension == ".gif":
        return header.startswith((b"GIF87a", b"GIF89a"))
    if extension == ".bmp":
        return header.startswith(b"BM")
    if extension == ".webp":
        return header.startswith(b"RIFF") and header[8:12] == b"WEBP"
    return False


async def _store_upload(
    file: UploadFile,
    *,
    user_id: int,
    category: str,
    allowed_extensions: set[str],
    max_bytes: int,
    require_image_signature: bool = False,
) -> dict:
    original_name = _safe_filename(file.filename or "")
    extension = Path(original_name).suffix.lower()
    if extension not in allowed_extensions:
        raise HTTPException(status_code=400, detail=f"File type {extension} not allowed")

    upload_dir = user_upload_root(user_id, category)
    upload_dir.mkdir(parents=True, exist_ok=True)
    destination = upload_dir / f"{uuid.uuid4().hex}_{original_name}"
    temporary = destination.with_suffix(destination.suffix + ".uploading")
    size = 0
    try:
        async with aiofiles.open(temporary, "wb") as output:
            while chunk := await file.read(_CHUNK_SIZE):
                size += len(chunk)
                if size > max_bytes:
                    raise HTTPException(
                        status_code=413,
                        detail=f"File too large (max {max_bytes // (1024 * 1024)}MB)",
                    )
                await output.write(chunk)
        if require_image_signature and not _looks_like_image(temporary, extension):
            raise HTTPException(status_code=400, detail="File content is not a valid image")
        os.replace(temporary, destination)
    except Exception:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
        raise
    finally:
        await file.close()

    return {"filename": original_name, "path": str(destination), "size": size}


@router.post("/attachment")
async def upload_attachment(
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user),
):
    return await _store_upload(
        file,
        user_id=current_user.id,
        category="attachments",
        allowed_extensions=ALLOWED_ATTACHMENTS,
        max_bytes=max(1, int(settings.MAX_ATTACHMENT_SIZE_MB)) * 1024 * 1024,
    )


@router.post("/image")
async def upload_image(
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user),
):
    return await _store_upload(
        file,
        user_id=current_user.id,
        category="images",
        allowed_extensions=ALLOWED_IMAGES,
        max_bytes=max(1, int(settings.MAX_IMAGE_SIZE_MB)) * 1024 * 1024,
        require_image_signature=True,
    )
