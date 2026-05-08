from datetime import datetime
from fastapi import APIRouter, Depends, HTTPException, Response
from fastapi.responses import RedirectResponse
from sqlmodel import Session
from ..database import get_session
from ..models.send_log import SendLog

router = APIRouter(tags=["追踪"])


@router.get("/track/open/{log_id}")
def track_open(log_id: int, session: Session = Depends(get_session)):
    log = session.get(SendLog, log_id)
    if not log:
        raise HTTPException(status_code=404, detail="Log not found")

    if not log.opened_at:
        log.opened_at = datetime.utcnow()
        session.add(log)
        session.commit()

    # Return 1x1 transparent GIF
    gif_data = b"\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff\x21\xf9\x04\x01\x00\x00\x00\x00\x2c\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02\x44\x01\x00\x3b"
    return Response(content=gif_data, media_type="image/gif")


@router.get("/track/click/{log_id}")
def track_click(log_id: int, url: str, session: Session = Depends(get_session)):
    log = session.get(SendLog, log_id)
    if not log:
        raise HTTPException(status_code=404, detail="Log not found")

    if not log.clicked_at:
        log.clicked_at = datetime.utcnow()
        session.add(log)
        session.commit()

    # Validate URL (basic check to prevent open redirect)
    if not url.startswith(("http://", "https://")):
        raise HTTPException(status_code=400, detail="Invalid URL")

    # 302 redirect to original URL
    return RedirectResponse(url=url, status_code=302)