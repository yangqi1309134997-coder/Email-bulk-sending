from datetime import datetime
from typing import Optional
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from sqlmodel import Session, select
from .deps import get_current_user
from ..database import get_session
from ..models.template import Template
from ..models.user import User

router = APIRouter(prefix="/api/templates", tags=["模板"])


class TemplateCreate(BaseModel):
    name: str
    subject: str = ""
    body: str = ""
    variables: list[str] = []


class TemplateUpdate(BaseModel):
    name: Optional[str] = None
    subject: Optional[str] = None
    body: Optional[str] = None
    variables: Optional[list[str]] = None


class TemplateResponse(BaseModel):
    id: int
    name: str
    subject: str
    body: str
    variables: str
    created_at: datetime
    updated_at: datetime

    class Config:
        from_attributes = True


@router.get("", response_model=list[TemplateResponse])
def list_templates(current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    templates = session.exec(select(Template).where(Template.user_id == current_user.id)).all()
    return templates


@router.post("", response_model=TemplateResponse)
def create_template(req: TemplateCreate, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    template = Template(
        user_id=current_user.id,
        name=req.name,
        subject=req.subject,
        body=req.body,
        variables=str(req.variables),
    )
    session.add(template)
    session.commit()
    session.refresh(template)
    return template


@router.put("/{template_id}", response_model=TemplateResponse)
def update_template(template_id: int, req: TemplateUpdate, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    template = session.get(Template, template_id)
    if not template or template.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Template not found")

    update_data = req.model_dump(exclude_unset=True)
    if "variables" in update_data:
        update_data["variables"] = str(update_data.pop("variables"))

    for key, value in update_data.items():
        setattr(template, key, value)

    template.updated_at = datetime.utcnow()
    session.add(template)
    session.commit()
    session.refresh(template)
    return template


@router.delete("/{template_id}")
def delete_template(template_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    template = session.get(Template, template_id)
    if not template or template.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Template not found")
    session.delete(template)
    session.commit()
    return {"message": "Template deleted"}


@router.post("/{template_id}/duplicate", response_model=TemplateResponse)
def duplicate_template(template_id: int, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    template = session.get(Template, template_id)
    if not template or template.user_id != current_user.id:
        raise HTTPException(status_code=404, detail="Template not found")

    new_template = Template(
        user_id=current_user.id,
        name=f"{template.name} (副本)",
        subject=template.subject,
        body=template.body,
        variables=template.variables,
    )
    session.add(new_template)
    session.commit()
    session.refresh(new_template)
    return new_template