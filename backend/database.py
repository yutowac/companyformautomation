"""SQLAlchemy engine, session, and startup init."""
from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime, timezone

from sqlalchemy import create_engine
from sqlalchemy.orm import declarative_base, sessionmaker

from config import (
    AUTH_TEST_EMAIL,
    AUTH_TEST_PASSWORD,
    DATABASE_URL,
    SAMPLE_APPLIED_EMAIL,
    SAMPLE_APPLIED_PASSWORD,
    SEED_SAMPLE_DATA,
)

Base = declarative_base()

_engine = None
SessionLocal = None


def _normalize_database_url(url: str) -> str:
    if url.startswith("postgres://"):
        return url.replace("postgres://", "postgresql://", 1)
    return url


def _ensure_engine():
    global _engine, SessionLocal
    if SessionLocal is not None:
        return
    if not DATABASE_URL:
        return
    url = _normalize_database_url(DATABASE_URL.strip())
    _engine = create_engine(url, pool_pre_ping=True)
    SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=_engine)


def get_engine():
    _ensure_engine()
    return _engine


def get_session_local():
    _ensure_engine()
    return SessionLocal


def init_db() -> None:
    """Create tables and seed test / sample users when DATABASE_URL is set."""
    _ensure_engine()
    if SessionLocal is None:
        print("DATABASE_URL not set; skipping DB table creation and auth seed.")
        return
    # Import models so tables are registered on Base.metadata
    import models  # noqa: F401

    engine = get_engine()
    assert engine is not None
    Base.metadata.create_all(bind=engine)
    _seed_test_user()
    if SEED_SAMPLE_DATA:
        _seed_sample_data()


def _seed_test_user() -> None:
    from auth_service import hash_password
    from models import User

    SL = get_session_local()
    assert SL is not None
    with SL() as db:
        existing = db.query(User).filter(User.email == AUTH_TEST_EMAIL).first()
        if existing:
            return
        user = User(
            email=AUTH_TEST_EMAIL,
            password_hash=hash_password(AUTH_TEST_PASSWORD),
        )
        db.add(user)
        db.commit()
        print(f"Seeded test user: {AUTH_TEST_EMAIL}")


def _sample_application_payload(email: str) -> dict:
    """FormData-shaped JSON for the applied sample user."""
    return {
        "companyName": "Sample Local Co",
        "presidentName": "Taro Yamada",
        "presidentNameLocal": "山田 太郎",
        "presidentAddress": "1-2-3 Chiyoda, Tokyo",
        "presidentAddressLocal": "東京都千代田区1-2-3",
        "birthyear": 1990,
        "birthmonth": 4,
        "birthday": 15,
        "purpose1": "IT consulting business",
        "purpose2": "Software development",
        "purpose3": "",
        "purpose4": "",
        "purpose5": "",
        "email": email,
    }


def _seed_sample_data() -> None:
    """Seed applied user + one pending application (local / SEED_SAMPLE_DATA only)."""
    from auth_service import hash_password
    from models import Application, ApplicationStatus, User

    SL = get_session_local()
    assert SL is not None
    with SL() as db:
        user = db.query(User).filter(User.email == SAMPLE_APPLIED_EMAIL).first()
        if not user:
            user = User(
                email=SAMPLE_APPLIED_EMAIL,
                password_hash=hash_password(SAMPLE_APPLIED_PASSWORD),
            )
            db.add(user)
            db.flush()
            print(f"Seeded sample applied user: {SAMPLE_APPLIED_EMAIL}")

        existing_app = (
            db.query(Application).filter(Application.user_id == user.id).first()
        )
        if existing_app:
            db.commit()
            return

        now = datetime.now(timezone.utc)
        app = Application(
            user_id=user.id,
            status=ApplicationStatus.PENDING.value,
            payload=_sample_application_payload(SAMPLE_APPLIED_EMAIL),
            submitted_at=now,
            updated_at=now,
        )
        db.add(app)
        db.commit()
        print(
            f"Seeded sample application for {SAMPLE_APPLIED_EMAIL} "
            f"(status={ApplicationStatus.PENDING.value})"
        )


def get_db():
    """FastAPI dependency: DB session."""
    from fastapi import HTTPException

    SL = get_session_local()
    if SL is None:
        raise HTTPException(status_code=503, detail="Database is not configured (set DATABASE_URL)")
    db = SL()
    try:
        yield db
    finally:
        db.close()


@contextmanager
def session_scope():
    """Non-FastAPI callers (e.g. tests)."""
    SL = get_session_local()
    if SL is None:
        raise RuntimeError("DATABASE_URL is not set")
    db = SL()
    try:
        yield db
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()
