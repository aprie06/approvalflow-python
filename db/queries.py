"""
db/queries.py

Query-layer functions built on the models in db/models.py. This is where the
VBA system's core operations get reimplemented against real foreign keys
instead of string-based lookups and fuzzy matching.

Notably, FindBestMatchForReply in the original VBA module used a scoring
system (supervisor email match, CC email match, name-in-subject match,
name-in-body match, time proximity) because Excel rows have no real foreign
keys. Here, a reply is matched to its originating SentLog row directly via
sent_log_id, a real foreign key. The fuzzy scoring logic is not ported,
it is made unnecessary by the schema.
"""

from datetime import datetime, date
from typing import Optional

from sqlalchemy.orm import Session
from sqlalchemy.dialects.postgresql import insert as pg_insert
from sqlalchemy import select

from db.models import (
    Organization,
    Supervisor,
    Intern,
    PayPeriod,
    SentLog,
    ReplyLog,
    SubmissionTracking,
)


# ---------------------------------------------------------------------------
# Lookups
# Replaces LoadLookupDataIntoMemory / GetStudentName / GetSupervisor / GetEmployer
# ---------------------------------------------------------------------------

def get_intern_by_email(session: Session, email: str) -> Optional[Intern]:
    """Looks up an intern by primary or student (CC) email, case-insensitive."""
    email = email.strip().lower()
    return session.scalar(
        select(Intern).where(
            (Intern.email.ilike(email)) | (Intern.student_email.ilike(email))
        )
    )


def get_supervisor_by_email(session: Session, email: str) -> Optional[Supervisor]:
    email = email.strip().lower()
    return session.scalar(select(Supervisor).where(Supervisor.email.ilike(email)))


# ---------------------------------------------------------------------------
# Pay periods
# Replaces GetPayPeriodStartDate / GetPayPeriodEndDate / the J1 anchor cell
# ---------------------------------------------------------------------------

def get_current_pay_period(session: Session, reference_date: date) -> Optional[PayPeriod]:
    """Returns the pay period containing reference_date, or None if not yet created."""
    return session.scalar(
        select(PayPeriod).where(
            PayPeriod.start_date <= reference_date,
            PayPeriod.end_date >= reference_date,
        )
    )


def get_or_create_pay_period(
    session: Session,
    label: str,
    start_date: date,
    end_date: date,
    submission_deadline: date,
    payroll_deadline: date,
) -> PayPeriod:
    """
    Idempotent: safe to call every time a new period opens, mirrors
    StartNewPayPeriodComplete without the "clear logs" side effects, since
    those are now unnecessary — each period is its own set of rows, not a
    workbook that gets wiped and reused.
    """
    existing = session.scalar(
        select(PayPeriod).where(
            PayPeriod.start_date == start_date, PayPeriod.end_date == end_date
        )
    )
    if existing:
        return existing

    period = PayPeriod(
        label=label,
        start_date=start_date,
        end_date=end_date,
        submission_deadline=submission_deadline,
        payroll_deadline=payroll_deadline,
    )
    session.add(period)
    session.flush()  # populate period.id without committing
    return period


# ---------------------------------------------------------------------------
# Sent log
# Replaces MoveItems' logging call (excel_log) and FastRebuildSentLog
# ---------------------------------------------------------------------------

def log_sent_email(
    session: Session,
    intern_id: int,
    supervisor_id: int,
    pay_period_id: int,
    sent_at: datetime,
    email_subject: Optional[str] = None,
    hours_reported: Optional[float] = None,
    submission_date: Optional[date] = None,
) -> SentLog:
    """One row per forwarded approval email. No dedup needed here by design,
    a legitimate resubmission is a legitimate second row. If accidental
    double-sends turn out to be a real problem in production, add a
    partial unique index on (intern_id, pay_period_id, sent_at) rounded
    to the minute, rather than deduplicating after the fact the way
    DeduplicateSentLog had to."""
    entry = SentLog(
        intern_id=intern_id,
        supervisor_id=supervisor_id,
        pay_period_id=pay_period_id,
        sent_at=sent_at,
        email_subject=email_subject,
        hours_reported=hours_reported,
        submission_date=submission_date,
    )
    session.add(entry)
    session.flush()
    return entry


# ---------------------------------------------------------------------------
# Reply log
# Replaces LogSupervisorReply / FindBestMatchForReply
# ---------------------------------------------------------------------------

def log_supervisor_reply(
    session: Session,
    sent_log_id: int,
    intern_id: int,
    supervisor_id: int,
    pay_period_id: int,
    received_at: datetime,
    response_type: str,
    match_method: str = "sent_log_fk",
    reply_body: Optional[str] = None,
) -> ReplyLog:
    """
    match_method defaults to 'sent_log_fk' since the caller is expected to
    have already identified sent_log_id directly (e.g. by parsing a reply
    threaded to a specific email, or by CC-matching the intern's address
    and querying the most recent open SentLog row for that intern/period).
    Use 'manual' if a human resolved an ambiguous case by hand.
    """
    if response_type not in ("APPROVED", "REJECTED", "CORRECTIONS"):
        raise ValueError(f"Invalid response_type: {response_type!r}")

    entry = ReplyLog(
        sent_log_id=sent_log_id,
        intern_id=intern_id,
        supervisor_id=supervisor_id,
        pay_period_id=pay_period_id,
        received_at=received_at,
        response_type=response_type,
        match_method=match_method,
        reply_body=reply_body,
    )
    session.add(entry)
    session.flush()
    return entry


def find_open_sent_log_for_intern(
    session: Session, intern_id: int, pay_period_id: int
) -> Optional[SentLog]:
    """
    Finds the most recent SentLog row for this intern in this pay period
    that does not yet have a reply. This is the direct-FK replacement for
    the CC-matching branch of FindBestMatchForReply: once you know the
    intern (from a CC address, a reply thread, or wherever), there is no
    scoring needed, just the most recent unanswered submission.
    """
    replied_sent_log_ids = select(ReplyLog.sent_log_id).where(
        ReplyLog.sent_log_id.is_not(None)
    )
    return session.scalar(
        select(SentLog)
        .where(
            SentLog.intern_id == intern_id,
            SentLog.pay_period_id == pay_period_id,
            SentLog.id.not_in(replied_sent_log_ids),
        )
        .order_by(SentLog.sent_at.desc())
    )


# ---------------------------------------------------------------------------
# Submission tracking
# Replaces PopulateSubmissionTracking / UpdateSubmissionStatus
#
# This is the direct fix for the VBA system's worst bug: the silent append
# that caused 2x row overcounting. Postgres's ON CONFLICT DO UPDATE, backed
# by the UNIQUE (pay_period_id, intern_id) constraint in schema.sql, makes
# "clear then repopulate" unnecessary — there is structurally no way to end
# up with two rows for the same intern in the same pay period.
# ---------------------------------------------------------------------------

def upsert_submission_tracking(
    session: Session,
    pay_period_id: int,
    intern_id: int,
    submitted_at: Optional[datetime] = None,
    approved_at: Optional[datetime] = None,
    status: str = "PENDING",
    notes: Optional[str] = None,
) -> None:
    if status not in ("PENDING", "APPROVED", "REJECTED", "CORRECTIONS", "MISSING"):
        raise ValueError(f"Invalid status: {status!r}")

    stmt = pg_insert(SubmissionTracking).values(
        pay_period_id=pay_period_id,
        intern_id=intern_id,
        submitted_at=submitted_at,
        approved_at=approved_at,
        status=status,
        notes=notes,
    )
    stmt = stmt.on_conflict_do_update(
        index_elements=["pay_period_id", "intern_id"],
        set_={
            "submitted_at": stmt.excluded.submitted_at,
            "approved_at": stmt.excluded.approved_at,
            "status": stmt.excluded.status,
            "notes": stmt.excluded.notes,
            "updated_at": func_now(),
        },
    )
    session.execute(stmt)


def func_now():
    """Small helper so the module only imports sqlalchemy.func once, here."""
    from sqlalchemy import func
    return func.now()


def populate_submission_tracking_for_period(
    session: Session, pay_period_id: int, intern_ids: list[int]
) -> int:
    """
    Ensures a PENDING row exists for every intern in intern_ids for this
    pay period, without disturbing rows that already have real status.
    Replaces PopulateSubmissionTracking. Safe to call repeatedly, since the
    upsert only sets status back to PENDING for interns that don't already
    have a row; existing rows are left as they are unless explicitly updated
    elsewhere in the workflow.
    """
    created = 0
    for intern_id in intern_ids:
        existing = session.scalar(
            select(SubmissionTracking).where(
                SubmissionTracking.pay_period_id == pay_period_id,
                SubmissionTracking.intern_id == intern_id,
            )
        )
        if existing is None:
            session.add(
                SubmissionTracking(
                    pay_period_id=pay_period_id,
                    intern_id=intern_id,
                    status="PENDING",
                )
            )
            created += 1
    session.flush()
    return created
