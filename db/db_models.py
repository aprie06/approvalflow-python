"""
db/models.py

SQLAlchemy ORM models for ApprovalFlow, mirroring schema.sql exactly.
Replaces Student_Sup_email.xlsx and all its sheets from the original VBA system.

Connection is read from the DATABASE_URL environment variable. Locally and in
CI this points at a disposable Postgres instance (see pytest.ini / the GitHub
Actions workflow). In production it should be set to the Azure PostgreSQL
Flexible Server connection string.
"""

import os
from datetime import datetime, date
from typing import Optional, List

from sqlalchemy import (
    create_engine,
    ForeignKey,
    String,
    Text,
    Numeric,
    Boolean,
    Date,
    DateTime,
    CheckConstraint,
    UniqueConstraint,
    Index,
    func,
)
from sqlalchemy.orm import (
    DeclarativeBase,
    Mapped,
    mapped_column,
    relationship,
    sessionmaker,
)


class Base(DeclarativeBase):
    pass


class Organization(Base):
    __tablename__ = "organizations"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(255), nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now())

    supervisors: Mapped[List["Supervisor"]] = relationship(back_populates="organization")
    interns: Mapped[List["Intern"]] = relationship(back_populates="organization")

    def __repr__(self) -> str:
        return f"<Organization id={self.id} name={self.name!r}>"


class Supervisor(Base):
    __tablename__ = "supervisors"

    id: Mapped[int] = mapped_column(primary_key=True)
    first_name: Mapped[str] = mapped_column(String(100), nullable=False)
    last_name: Mapped[str] = mapped_column(String(100), nullable=False)
    email: Mapped[str] = mapped_column(String(255), nullable=False, unique=True)
    organization_id: Mapped[Optional[int]] = mapped_column(ForeignKey("organizations.id"))
    active: Mapped[bool] = mapped_column(Boolean, default=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now())

    organization: Mapped[Optional["Organization"]] = relationship(back_populates="supervisors")
    interns: Mapped[List["Intern"]] = relationship(back_populates="supervisor")

    def __repr__(self) -> str:
        return f"<Supervisor id={self.id} email={self.email!r}>"


class Intern(Base):
    __tablename__ = "interns"

    id: Mapped[int] = mapped_column(primary_key=True)
    first_name: Mapped[str] = mapped_column(String(100), nullable=False)
    last_name: Mapped[str] = mapped_column(String(100), nullable=False)
    email: Mapped[str] = mapped_column(String(255), nullable=False, unique=True)
    student_email: Mapped[Optional[str]] = mapped_column(String(255))  # CC-only address
    employee_id: Mapped[Optional[str]] = mapped_column(String(50))
    supervisor_id: Mapped[Optional[int]] = mapped_column(ForeignKey("supervisors.id"))
    organization_id: Mapped[Optional[int]] = mapped_column(ForeignKey("organizations.id"))
    start_date: Mapped[Optional[date]] = mapped_column(Date)
    active: Mapped[bool] = mapped_column(Boolean, default=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now())

    supervisor: Mapped[Optional["Supervisor"]] = relationship(back_populates="interns")
    organization: Mapped[Optional["Organization"]] = relationship(back_populates="interns")

    __table_args__ = (
        Index("idx_interns_email", "email"),
    )

    def __repr__(self) -> str:
        return f"<Intern id={self.id} email={self.email!r}>"


class PayPeriod(Base):
    __tablename__ = "pay_periods"

    id: Mapped[int] = mapped_column(primary_key=True)
    label: Mapped[str] = mapped_column(String(50), nullable=False)
    start_date: Mapped[date] = mapped_column(Date, nullable=False)
    end_date: Mapped[date] = mapped_column(Date, nullable=False)
    submission_deadline: Mapped[date] = mapped_column(Date, nullable=False)
    payroll_deadline: Mapped[date] = mapped_column(Date, nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now())

    __table_args__ = (
        UniqueConstraint("start_date", "end_date", name="uq_pay_periods_start_end"),
    )

    def __repr__(self) -> str:
        return f"<PayPeriod id={self.id} label={self.label!r}>"


class SentLog(Base):
    """One row per forwarded approval email. Replaces the Sent_Log sheet."""

    __tablename__ = "sent_log"

    id: Mapped[int] = mapped_column(primary_key=True)
    pay_period_id: Mapped[Optional[int]] = mapped_column(ForeignKey("pay_periods.id"))
    intern_id: Mapped[Optional[int]] = mapped_column(ForeignKey("interns.id"))
    supervisor_id: Mapped[Optional[int]] = mapped_column(ForeignKey("supervisors.id"))
    sent_at: Mapped[datetime] = mapped_column(DateTime, nullable=False)
    email_subject: Mapped[Optional[str]] = mapped_column(String(500))
    hours_reported: Mapped[Optional[float]] = mapped_column(Numeric(5, 2))
    submission_date: Mapped[Optional[date]] = mapped_column(Date)
    created_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now())

    __table_args__ = (
        Index("idx_sent_log_pay_period", "pay_period_id"),
    )

    def __repr__(self) -> str:
        return f"<SentLog id={self.id} intern_id={self.intern_id} sent_at={self.sent_at}>"


class ReplyLog(Base):
    """One row per parsed supervisor reply. Replaces the Reply_Log sheet."""

    __tablename__ = "reply_log"

    id: Mapped[int] = mapped_column(primary_key=True)
    sent_log_id: Mapped[Optional[int]] = mapped_column(ForeignKey("sent_log.id"))
    pay_period_id: Mapped[Optional[int]] = mapped_column(ForeignKey("pay_periods.id"))
    intern_id: Mapped[Optional[int]] = mapped_column(ForeignKey("interns.id"))
    supervisor_id: Mapped[Optional[int]] = mapped_column(ForeignKey("supervisors.id"))
    received_at: Mapped[datetime] = mapped_column(DateTime, nullable=False)
    response_type: Mapped[Optional[str]] = mapped_column(String(20))
    match_method: Mapped[Optional[str]] = mapped_column(String(50))  # 'cc_match', 'name_score', 'manual'
    reply_body: Mapped[Optional[str]] = mapped_column(Text)
    created_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now())

    __table_args__ = (
        CheckConstraint(
            "response_type IN ('APPROVED', 'REJECTED', 'CORRECTIONS')",
            name="ck_reply_log_response_type",
        ),
    )

    def __repr__(self) -> str:
        return f"<ReplyLog id={self.id} response_type={self.response_type!r}>"


class SubmissionTracking(Base):
    """
    One row per intern per pay period. Replaces the Submission_Tracking sheet.

    The UNIQUE constraint on (pay_period_id, intern_id) is a deliberate fix for
    the VBA system's most dangerous bug: PopulateSubmissionTracking silently
    appended instead of replacing, causing 2x row overcounting in production.
    Here, a duplicate insert raises an IntegrityError instead of corrupting data.
    Always upsert against this constraint, never blind-insert.
    """

    __tablename__ = "submission_tracking"

    id: Mapped[int] = mapped_column(primary_key=True)
    pay_period_id: Mapped[Optional[int]] = mapped_column(ForeignKey("pay_periods.id"))
    intern_id: Mapped[Optional[int]] = mapped_column(ForeignKey("interns.id"))
    submitted_at: Mapped[Optional[datetime]] = mapped_column(DateTime)
    approved_at: Mapped[Optional[datetime]] = mapped_column(DateTime)
    status: Mapped[str] = mapped_column(String(20), default="PENDING")
    notes: Mapped[Optional[str]] = mapped_column(Text)
    created_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(
        DateTime, server_default=func.now(), onupdate=func.now()
    )

    __table_args__ = (
        UniqueConstraint("pay_period_id", "intern_id", name="uq_submission_tracking_period_intern"),
        CheckConstraint(
            "status IN ('PENDING', 'APPROVED', 'REJECTED', 'CORRECTIONS', 'MISSING')",
            name="ck_submission_tracking_status",
        ),
        Index("idx_submission_tracking_pay_period", "pay_period_id"),
    )

    def __repr__(self) -> str:
        return f"<SubmissionTracking id={self.id} status={self.status!r}>"


# ---------------------------------------------------------------------------
# Engine / session setup
# ---------------------------------------------------------------------------

def get_database_url() -> str:
    """
    Resolves the database connection string.

    DATABASE_URL is the primary variable, used locally and in CI against a
    disposable Postgres container. AZURE_POSTGRESQL_CONNECTION_STRING is
    checked as a fallback for production/Azure deployment, so the same
    codebase works in both environments without a code change.
    """
    url = os.environ.get("DATABASE_URL") or os.environ.get("AZURE_POSTGRESQL_CONNECTION_STRING")
    if not url:
        raise RuntimeError(
            "No database connection string found. Set DATABASE_URL (local/CI) "
            "or AZURE_POSTGRESQL_CONNECTION_STRING (production)."
        )
    return url


def get_engine():
    return create_engine(get_database_url(), pool_pre_ping=True)


def get_session_factory():
    return sessionmaker(bind=get_engine(), expire_on_commit=False)


def init_db() -> None:
    """Creates all tables if they do not already exist. Mirrors schema.sql."""
    Base.metadata.create_all(get_engine())


if __name__ == "__main__":
    init_db()
    print("Tables created (or already present).")
