from fastapi import APIRouter, Depends, HTTPException, Response
from fastapi.responses import RedirectResponse
from sqlalchemy import func, update
from sqlmodel import Session
from ..database import get_session
from ..models.send_log import SendLog
from ..models.task import Task
from ..services.tracker import is_valid_redirect_url
from ..utils.security import verify_tracking_signature
from ..utils.time import utcnow

router = APIRouter(tags=["追踪"])


@router.get("/track/open/{log_id}")
def track_open(log_id: int, sig: str = "", session: Session = Depends(get_session)):
    if not verify_tracking_signature(log_id, "open", sig):
        raise HTTPException(status_code=403, detail="Invalid tracking signature")
    log = session.get(SendLog, log_id)
    if not log:
        raise HTTPException(status_code=404, detail="Log not found")

    if not log.opened_at:
        opened_at = utcnow()
        result = session.exec(
            update(SendLog)
            .where(SendLog.id == log_id, SendLog.opened_at == None)  # noqa: E711
            .values(opened_at=opened_at)
            .execution_options(synchronize_session=False)
        )
        if result.rowcount == 1:
            session.exec(
                update(Task)
                .where(Task.id == log.task_id)
                .values(open_count=func.coalesce(Task.open_count, 0) + 1)
                .execution_options(synchronize_session=False)
            )
            session.commit()
        else:
            session.rollback()

    # Return 1x1 transparent GIF
    gif_data = b"\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff\x21\xf9\x04\x01\x00\x00\x00\x00\x2c\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02\x44\x01\x00\x3b"
    return Response(
        content=gif_data,
        media_type="image/gif",
        headers={"Cache-Control": "no-store, max-age=0"},
    )


@router.get("/track/click/{log_id}")
def track_click(log_id: int, url: str, sig: str = "", session: Session = Depends(get_session)):
    if not verify_tracking_signature(log_id, "click", sig):
        raise HTTPException(status_code=403, detail="Invalid tracking signature")
    if not is_valid_redirect_url(url):
        raise HTTPException(status_code=400, detail="Invalid redirect URL")

    log = session.get(SendLog, log_id)
    if not log:
        raise HTTPException(status_code=404, detail="Log not found")

    if not log.clicked_at:
        clicked_at = utcnow()
        result = session.exec(
            update(SendLog)
            .where(SendLog.id == log_id, SendLog.clicked_at == None)  # noqa: E711
            .values(clicked_at=clicked_at)
            .execution_options(synchronize_session=False)
        )
        if result.rowcount == 1:
            session.exec(
                update(Task)
                .where(Task.id == log.task_id)
                .values(click_count=func.coalesce(Task.click_count, 0) + 1)
                .execution_options(synchronize_session=False)
            )
            session.commit()
        else:
            session.rollback()

    # 302 redirect to original URL
    return RedirectResponse(url=url, status_code=302)
