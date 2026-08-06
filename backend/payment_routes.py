"""Monthly Payment Requests API routes."""
from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from fastapi import APIRouter, Depends, File, Form, Header, HTTPException, UploadFile, status
from pydantic import BaseModel
from sqlalchemy.orm import Session

from config import ADMIN_API_KEY, PAYMENT_REQUEST_MAX_ACTIVE
from database import get_db
from deps import get_current_user
from models import PaymentRequest, PaymentRequestStatus, User
from payment_service import (
    append_payment_to_spreadsheet,
    is_payment_editable_now,
    upload_payment_attachment,
    validate_payment_payload,
)

router = APIRouter(tags=["monthly-payment-requests"])


class PaymentRequestItem(BaseModel):
    id: int
    status: str
    payload: dict
    attachment_url: str | None
    attachment_name: str | None
    created_at: str
    submitted_at: str | None
    updated_at: str | None
    editable: bool


class PaymentListResponse(BaseModel):
    items: list[PaymentRequestItem]
    editable_window: bool
    remaining_slots: int
    max_active: int


def _iso(dt: datetime | None) -> str | None:
    return dt.isoformat() if dt else None


def _active_query(db: Session, user_id: int):
    return (
        db.query(PaymentRequest)
        .filter(
            PaymentRequest.user_id == user_id,
            PaymentRequest.status != PaymentRequestStatus.PAID.value,
        )
        .order_by(PaymentRequest.created_at.desc())
    )


def _to_item(row: PaymentRequest, editable_window: bool) -> PaymentRequestItem:
    return PaymentRequestItem(
        id=row.id,
        status=row.status,
        payload=row.payload or {},
        attachment_url=row.attachment_url,
        attachment_name=row.attachment_name,
        created_at=_iso(row.created_at) or "",
        submitted_at=_iso(row.submitted_at),
        updated_at=_iso(row.updated_at),
        editable=editable_window and row.status != PaymentRequestStatus.PAID.value,
    )


def _parse_payload_from_form(
    payeeName: str,
    bankName: str,
    branchName: str,
    accountType: str,
    accountNumber: str,
    accountHolderKana: str,
    amountJpy: str,
    invoiceNumber: str = "",
) -> dict[str, Any]:
    try:
        return validate_payment_payload(
            {
                "payeeName": payeeName,
                "bankName": bankName,
                "branchName": branchName,
                "accountType": accountType,
                "accountNumber": accountNumber,
                "accountHolderKana": accountHolderKana,
                "amountJpy": amountJpy,
                "invoiceNumber": invoiceNumber,
            }
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e


async def _maybe_upload(attachment: UploadFile | None) -> tuple[str, str]:
    if attachment is None or not attachment.filename:
        return "", ""
    content_type = attachment.content_type or ""
    if not content_type.startswith("image/"):
        raise HTTPException(status_code=400, detail="Attachment must be an image")
    data = await attachment.read()
    if not data:
        return "", ""
    if len(data) > 8 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="Attachment must be 8MB or smaller")
    url = upload_payment_attachment(data, attachment.filename, content_type)
    return url, attachment.filename


@router.get("/monthly-payment-requests", response_model=PaymentListResponse)
def list_payment_requests(
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    editable = is_payment_editable_now()
    rows = _active_query(db, current_user.id).all()
    max_active = PAYMENT_REQUEST_MAX_ACTIVE
    return PaymentListResponse(
        items=[_to_item(r, editable) for r in rows],
        editable_window=editable,
        remaining_slots=max(0, max_active - len(rows)),
        max_active=max_active,
    )


@router.get("/monthly-payment-requests/{request_id}", response_model=PaymentRequestItem)
def get_payment_request(
    request_id: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    row = (
        db.query(PaymentRequest)
        .filter(PaymentRequest.id == request_id, PaymentRequest.user_id == current_user.id)
        .first()
    )
    if not row or row.status == PaymentRequestStatus.PAID.value:
        raise HTTPException(status_code=404, detail="Payment request not found")
    return _to_item(row, is_payment_editable_now())


@router.post("/monthly-payment-requests", response_model=PaymentRequestItem)
async def create_payment_request(
    payeeName: str = Form(...),
    bankName: str = Form(...),
    branchName: str = Form(...),
    accountType: str = Form(...),
    accountNumber: str = Form(...),
    accountHolderKana: str = Form(...),
    amountJpy: str = Form(...),
    invoiceNumber: str = Form(""),
    attachment: UploadFile | None = File(None),
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    if not is_payment_editable_now():
        raise HTTPException(
            status_code=403,
            detail="Payment requests are frozen until the next editing window (1st–20th JST).",
        )
    active_count = _active_query(db, current_user.id).count()
    if active_count >= PAYMENT_REQUEST_MAX_ACTIVE:
        raise HTTPException(
            status_code=409,
            detail=f"Maximum of {PAYMENT_REQUEST_MAX_ACTIVE} active payment requests reached",
        )

    payload = _parse_payload_from_form(
        payeeName,
        bankName,
        branchName,
        accountType,
        accountNumber,
        accountHolderKana,
        amountJpy,
        invoiceNumber,
    )
    att_url, att_name = await _maybe_upload(attachment)
    now = datetime.now(timezone.utc)
    row = PaymentRequest(
        user_id=current_user.id,
        status=PaymentRequestStatus.PENDING.value,
        payload=payload,
        attachment_url=att_url or None,
        attachment_name=att_name or None,
        submitted_at=now,
        updated_at=now,
    )
    db.add(row)
    db.commit()
    db.refresh(row)

    append_payment_to_spreadsheet(
        user_email=current_user.email,
        payload=payload,
        request_id=row.id,
        status=row.status,
        attachment_url=att_url,
    )
    return _to_item(row, True)


@router.put("/monthly-payment-requests/{request_id}", response_model=PaymentRequestItem)
async def update_payment_request(
    request_id: int,
    payeeName: str = Form(...),
    bankName: str = Form(...),
    branchName: str = Form(...),
    accountType: str = Form(...),
    accountNumber: str = Form(...),
    accountHolderKana: str = Form(...),
    amountJpy: str = Form(...),
    invoiceNumber: str = Form(""),
    attachment: UploadFile | None = File(None),
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    if not is_payment_editable_now():
        raise HTTPException(
            status_code=403,
            detail="Payment requests are frozen until the next editing window (1st–20th JST).",
        )
    row = (
        db.query(PaymentRequest)
        .filter(PaymentRequest.id == request_id, PaymentRequest.user_id == current_user.id)
        .first()
    )
    if not row or row.status == PaymentRequestStatus.PAID.value:
        raise HTTPException(status_code=404, detail="Payment request not found")

    payload = _parse_payload_from_form(
        payeeName,
        bankName,
        branchName,
        accountType,
        accountNumber,
        accountHolderKana,
        amountJpy,
        invoiceNumber,
    )
    att_url, att_name = await _maybe_upload(attachment)
    row.payload = payload
    row.updated_at = datetime.now(timezone.utc)
    if att_url:
        row.attachment_url = att_url
        row.attachment_name = att_name
    db.commit()
    db.refresh(row)

    append_payment_to_spreadsheet(
        user_email=current_user.email,
        payload=payload,
        request_id=row.id,
        status=row.status,
        attachment_url=row.attachment_url or "",
    )
    return _to_item(row, True)


@router.post("/admin/monthly-payment-requests/{request_id}/complete")
def admin_complete_payment_request(
    request_id: int,
    x_admin_key: str | None = Header(default=None, alias="X-Admin-Key"),
    db: Session = Depends(get_db),
):
    if not ADMIN_API_KEY or x_admin_key != ADMIN_API_KEY:
        raise HTTPException(status_code=401, detail="Unauthorized")
    row = db.query(PaymentRequest).filter(PaymentRequest.id == request_id).first()
    if not row:
        raise HTTPException(status_code=404, detail="Payment request not found")
    if row.status == PaymentRequestStatus.PAID.value:
        return {"message": "Already marked paid", "id": row.id}
    row.status = PaymentRequestStatus.PAID.value
    row.paid_at = datetime.now(timezone.utc)
    row.updated_at = row.paid_at
    db.commit()
    return {"message": "Payment request marked paid", "id": row.id}
