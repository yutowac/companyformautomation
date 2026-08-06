"""ORM models."""
from __future__ import annotations

from enum import Enum

from sqlalchemy import JSON, Column, DateTime, ForeignKey, Integer, String, func
from sqlalchemy.orm import relationship

from database import Base


class ApplicationStatus(str, Enum):
    PENDING = "pending"
    IN_REVIEW = "in_review"
    COMPLETED = "completed"
    REJECTED = "rejected"


class PaymentRequestStatus(str, Enum):
    PENDING = "pending"
    PAID = "paid"


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String(255), unique=True, nullable=False, index=True)
    password_hash = Column(String(255), nullable=False)
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    applications = relationship("Application", back_populates="user")
    payment_requests = relationship("PaymentRequest", back_populates="user")


class Application(Base):
    __tablename__ = "applications"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=False, unique=True, index=True)
    status = Column(
        String(32),
        nullable=False,
        default=ApplicationStatus.PENDING.value,
        server_default=ApplicationStatus.PENDING.value,
    )
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    submitted_at = Column(DateTime(timezone=True), nullable=True)
    updated_at = Column(
        DateTime(timezone=True),
        nullable=True,
        server_default=func.now(),
        onupdate=func.now(),
    )
    payload = Column(JSON, nullable=False)

    user = relationship("User", back_populates="applications")


class PaymentRequest(Base):
    __tablename__ = "payment_requests"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=False, index=True)
    status = Column(
        String(32),
        nullable=False,
        default=PaymentRequestStatus.PENDING.value,
        server_default=PaymentRequestStatus.PENDING.value,
        index=True,
    )
    payload = Column(JSON, nullable=False)
    attachment_url = Column(String(1024), nullable=True)
    attachment_name = Column(String(512), nullable=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    submitted_at = Column(DateTime(timezone=True), nullable=True)
    updated_at = Column(
        DateTime(timezone=True),
        nullable=True,
        server_default=func.now(),
        onupdate=func.now(),
    )
    paid_at = Column(DateTime(timezone=True), nullable=True)

    user = relationship("User", back_populates="payment_requests")
