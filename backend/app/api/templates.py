import json
from datetime import datetime
from typing import Optional
from fastapi import APIRouter, Depends, HTTPException, Query
from pydantic import BaseModel, ConfigDict, Field, field_validator
from sqlmodel import Session, select
from .deps import get_current_user
from ..database import get_session
from ..models.template import Template
from ..models.user import User
from ..config import settings
from ..utils.time import utcnow

router = APIRouter(prefix="/api/templates", tags=["模板"])


class TemplateCreate(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    name: str = Field(min_length=1, max_length=200)
    subject: str = Field(default="", max_length=998)
    body: str = Field(default="")
    variables: list[str] = Field(default_factory=list, max_length=100)

    @field_validator("subject")
    @classmethod
    def validate_subject(cls, value: str) -> str:
        if any(ord(char) < 32 or ord(char) == 127 for char in value):
            raise ValueError("Template subject cannot contain control characters")
        return value

    @field_validator("body")
    @classmethod
    def validate_body_size(cls, value: str) -> str:
        if len(value.encode("utf-8")) > int(settings.MAX_TASK_BODY_BYTES):
            raise ValueError("Template body exceeds system limit")
        return value

    @field_validator("variables")
    @classmethod
    def validate_variables(cls, value: list[str]) -> list[str]:
        cleaned = []
        for variable in value:
            item = variable.strip()
            if not item or len(item) > 100 or any(ord(char) < 32 for char in item):
                raise ValueError("Invalid template variable")
            cleaned.append(item)
        return list(dict.fromkeys(cleaned))


class TemplateUpdate(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    name: Optional[str] = Field(default=None, min_length=1, max_length=200)
    subject: Optional[str] = Field(default=None, max_length=998)
    body: Optional[str] = None
    variables: Optional[list[str]] = Field(default=None, max_length=100)

    @field_validator("subject")
    @classmethod
    def validate_subject(cls, value: Optional[str]) -> Optional[str]:
        if value is not None and any(ord(char) < 32 or ord(char) == 127 for char in value):
            raise ValueError("Template subject cannot contain control characters")
        return value

    @field_validator("body")
    @classmethod
    def validate_body_size(cls, value: Optional[str]) -> Optional[str]:
        if value is not None and len(value.encode("utf-8")) > int(settings.MAX_TASK_BODY_BYTES):
            raise ValueError("Template body exceeds system limit")
        return value

    @field_validator("variables")
    @classmethod
    def validate_variables(cls, value: Optional[list[str]]) -> Optional[list[str]]:
        if value is None:
            return None
        cleaned = []
        for variable in value:
            item = variable.strip()
            if not item or len(item) > 100 or any(ord(char) < 32 for char in item):
                raise ValueError("Invalid template variable")
            cleaned.append(item)
        return list(dict.fromkeys(cleaned))


class TemplateResponse(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    name: str
    subject: str
    body: str
    variables: str
    created_at: datetime
    updated_at: datetime

@router.get("", response_model=list[TemplateResponse])
def list_templates(
    skip: int = Query(default=0, ge=0),
    limit: int = Query(default=100, ge=1, le=500),
    current_user: User = Depends(get_current_user),
    session: Session = Depends(get_session),
):
    templates = session.exec(
        select(Template)
        .where(Template.user_id == current_user.id)
        .order_by(Template.updated_at.desc())
        .offset(skip)
        .limit(limit)
    ).all()
    return templates


@router.post("", response_model=TemplateResponse)
def create_template(req: TemplateCreate, current_user: User = Depends(get_current_user), session: Session = Depends(get_session)):
    template = Template(
        user_id=current_user.id,
        name=req.name.strip(),
        subject=req.subject,
        body=req.body,
        variables=json.dumps(req.variables, ensure_ascii=False),
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
        update_data["variables"] = json.dumps(
            update_data.pop("variables"), ensure_ascii=False
        )
    if "name" in update_data:
        update_data["name"] = update_data["name"].strip()

    for key, value in update_data.items():
        setattr(template, key, value)

    template.updated_at = utcnow()
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
        variables=template.variables or "[]",
    )
    session.add(new_template)
    session.commit()
    session.refresh(new_template)
    return new_template
