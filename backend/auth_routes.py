"""Authentication and application listing routes."""
from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException, status
from pydantic import BaseModel, EmailStr
from sqlalchemy.orm import Session

from auth_service import create_access_token, hash_password, verify_password
from database import get_db
from deps import authenticate_user, get_current_user
from models import Application, User

router = APIRouter(tags=["auth"])


class LoginRequest(BaseModel):
    email: EmailStr
    password: str


class TokenResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"


class ChangePasswordRequest(BaseModel):
    current_password: str
    new_password: str


class MeResponse(BaseModel):
    email: str
    has_application: bool
    application_status: str | None = None
    application_submitted_at: str | None = None


class ApplicationListItem(BaseModel):
    id: int
    created_at: str
    submitted_at: str | None
    status: str
    payload: dict


@router.post("/auth/login", response_model=TokenResponse)
def login(body: LoginRequest, db: Session = Depends(get_db)):
    user = authenticate_user(db, body.email, body.password)
    if not user:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect email or password",
        )
    token = create_access_token(user.id)
    return TokenResponse(access_token=token)


@router.post("/auth/change-password")
def change_password(
    body: ChangePasswordRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    if not body.new_password or len(body.new_password) < 8:
        raise HTTPException(status_code=400, detail="New password must be at least 8 characters")
    if not verify_password(body.current_password, current_user.password_hash):
        raise HTTPException(status_code=400, detail="Current password is incorrect")
    current_user.password_hash = hash_password(body.new_password)
    db.add(current_user)
    db.commit()
    return {"message": "Password updated"}


@router.get("/me", response_model=MeResponse)
def me(current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    row = (
        db.query(Application)
        .filter(Application.user_id == current_user.id)
        .order_by(Application.created_at.desc())
        .first()
    )
    if not row:
        return MeResponse(
            email=current_user.email,
            has_application=False,
            application_status=None,
            application_submitted_at=None,
        )
    submitted = row.submitted_at.isoformat() if row.submitted_at else None
    return MeResponse(
        email=current_user.email,
        has_application=True,
        application_status=row.status,
        application_submitted_at=submitted,
    )


@router.get("/applications", response_model=list[ApplicationListItem])
def list_applications(current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    rows = (
        db.query(Application)
        .filter(Application.user_id == current_user.id)
        .order_by(Application.created_at.desc())
        .all()
    )
    out: list[ApplicationListItem] = []
    for r in rows:
        created = r.created_at.isoformat() if r.created_at else ""
        submitted = r.submitted_at.isoformat() if r.submitted_at else None
        out.append(
            ApplicationListItem(
                id=r.id,
                created_at=created,
                submitted_at=submitted,
                status=r.status or "pending",
                payload=r.payload or {},
            )
        )
    return out
