"""
tests/test_queries.py

Tests for db/queries.py against a real Postgres database, pointed to by
DATABASE_URL. In CI this is the disposable service container defined in
.github/workflows/pipeline.yml. Locally, point it at the Docker Compose
instance from `docker compose up -d`.

Each test runs inside its own transaction that is rolled back afterward,
so tests do not interfere with each other and never need manual cleanup.
"""

import os
from datetime import date, datetime

import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from db.models import Base, Organization, Supervisor, Intern
from db.queries import (
    get_intern_by_email,
    get_supervisor_by_email,
    get_or_create_pay_period,
    get_current_pay_period,
    log_sent_email,
    find_open_sent_log_for_intern,
    log_supervisor_reply,
    upsert_submission_tracking,
    populate_submission_tracking_for_period,
)
from db.models import SubmissionTracking


@pytest.fixture(scope="session")
def engine():
    database_url = os.environ.get("DATABASE_URL")
    if not database_url:
        pytest.skip("DATABASE_URL not set, skipping database tests")
    eng = create_engine(database_url)
    Base.metadata.create_all(eng)
    return eng


@pytest.fixture
def db_session(engine):
    """
    Each test gets its own connection and transaction, rolled back at
    teardown. Nothing a test writes persists past that test, and tests
    never need to know or care about what other tests inserted.
    """
    connection = engine.connect()
    trans = connection.begin()
    Session = sessionmaker(bind=connection, expire_on_commit=False)
    session = Session()
    yield session
    session.close()
    trans.rollback()
    connection.close()


@pytest.fixture
def sample_org(db_session):
    org = Organization(name="Test Org")
    db_session.add(org)
    db_session.flush()
    return org


@pytest.fixture
def sample_supervisor(db_session, sample_org):
    sup = Supervisor(
        first_name="Jane",
        last_name="Doe",
        email="jane.doe@example.org",
        organization=sample_org,
    )
    db_session.add(sup)
    db_session.flush()
    return sup


@pytest.fixture
def sample_intern(db_session, sample_org, sample_supervisor):
    intern = Intern(
        first_name="Sam",
        last_name="Lee",
        email="sam.lee@example.org",
        student_email="samlee@student.example.edu",
        supervisor=sample_supervisor,
        organization=sample_org,
    )
    db_session.add(intern)
    db_session.flush()
    return intern


@pytest.fixture
def sample_pay_period(db_session):
    return get_or_create_pay_period(
        db_session,
        label="Jul 1-15 2026",
        start_date=date(2026, 7, 1),
        end_date=date(2026, 7, 15),
        submission_deadline=date(2026, 7, 15),
        payroll_deadline=date(2026, 7, 17),
    )


# ---------------------------------------------------------------------------
# Lookups
# ---------------------------------------------------------------------------

def test_get_intern_by_email_case_insensitive(db_session, sample_intern):
    found = get_intern_by_email(db_session, "SAM.LEE@EXAMPLE.ORG")
    assert found is not None
    assert found.id == sample_intern.id


def test_get_intern_by_email_matches_student_email(db_session, sample_intern):
    found = get_intern_by_email(db_session, "samlee@student.example.edu")
    assert found is not None
    assert found.id == sample_intern.id


def test_get_intern_by_email_no_match_returns_none(db_session, sample_intern):
    assert get_intern_by_email(db_session, "nobody@example.org") is None


def test_get_supervisor_by_email(db_session, sample_supervisor):
    found = get_supervisor_by_email(db_session, "jane.doe@example.org")
    assert found is not None
    assert found.id == sample_supervisor.id


# ---------------------------------------------------------------------------
# Pay periods
# ---------------------------------------------------------------------------

def test_get_or_create_pay_period_is_idempotent(db_session):
    p1 = get_or_create_pay_period(
        db_session, "Jun 16-30 2026",
        date(2026, 6, 16), date(2026, 6, 30),
        date(2026, 6, 30), date(2026, 7, 2),
    )
    p2 = get_or_create_pay_period(
        db_session, "Jun 16-30 2026 (duplicate call)",
        date(2026, 6, 16), date(2026, 6, 30),
        date(2026, 6, 30), date(2026, 7, 2),
    )
    assert p1.id == p2.id


def test_get_current_pay_period_finds_containing_period(db_session, sample_pay_period):
    found = get_current_pay_period(db_session, date(2026, 7, 8))
    assert found is not None
    assert found.id == sample_pay_period.id


def test_get_current_pay_period_returns_none_outside_range(db_session, sample_pay_period):
    assert get_current_pay_period(db_session, date(2026, 8, 1)) is None


# ---------------------------------------------------------------------------
# Sent log / reply log / open-submission matching
# ---------------------------------------------------------------------------

def test_log_sent_email_creates_row(db_session, sample_intern, sample_supervisor, sample_pay_period):
    sent = log_sent_email(
        db_session,
        intern_id=sample_intern.id,
        supervisor_id=sample_supervisor.id,
        pay_period_id=sample_pay_period.id,
        sent_at=datetime(2026, 7, 2, 9, 0),
        email_subject="Timesheet Approval Request for Sam Lee",
    )
    assert sent.id is not None
    assert sent.email_subject == "Timesheet Approval Request for Sam Lee"


def test_find_open_sent_log_before_and_after_reply(
    db_session, sample_intern, sample_supervisor, sample_pay_period
):
    sent = log_sent_email(
        db_session,
        intern_id=sample_intern.id,
        supervisor_id=sample_supervisor.id,
        pay_period_id=sample_pay_period.id,
        sent_at=datetime(2026, 7, 2, 9, 0),
    )

    open_before = find_open_sent_log_for_intern(db_session, sample_intern.id, sample_pay_period.id)
    assert open_before is not None
    assert open_before.id == sent.id

    log_supervisor_reply(
        db_session,
        sent_log_id=sent.id,
        intern_id=sample_intern.id,
        supervisor_id=sample_supervisor.id,
        pay_period_id=sample_pay_period.id,
        received_at=datetime(2026, 7, 2, 14, 0),
        response_type="APPROVED",
    )

    open_after = find_open_sent_log_for_intern(db_session, sample_intern.id, sample_pay_period.id)
    assert open_after is None


def test_log_supervisor_reply_rejects_invalid_response_type(
    db_session, sample_intern, sample_supervisor, sample_pay_period
):
    sent = log_sent_email(
        db_session,
        intern_id=sample_intern.id,
        supervisor_id=sample_supervisor.id,
        pay_period_id=sample_pay_period.id,
        sent_at=datetime(2026, 7, 2, 9, 0),
    )
    with pytest.raises(ValueError):
        log_supervisor_reply(
            db_session,
            sent_log_id=sent.id,
            intern_id=sample_intern.id,
            supervisor_id=sample_supervisor.id,
            pay_period_id=sample_pay_period.id,
            received_at=datetime(2026, 7, 2, 14, 0),
            response_type="MAYBE",
        )


# ---------------------------------------------------------------------------
# Submission tracking / upsert
#
# This is the direct regression test for the VBA system's worst bug: the
# silent append that caused 2x row overcounting. If this constraint or the
# upsert logic is ever broken, this test fails loudly instead of silently
# doubling rows the way PopulateSubmissionTracking did in production.
# ---------------------------------------------------------------------------

def test_upsert_submission_tracking_inserts_once(db_session, sample_intern, sample_pay_period):
    upsert_submission_tracking(
        db_session,
        pay_period_id=sample_pay_period.id,
        intern_id=sample_intern.id,
        status="PENDING",
    )
    db_session.flush()

    rows = (
        db_session.query(SubmissionTracking)
        .filter_by(pay_period_id=sample_pay_period.id, intern_id=sample_intern.id)
        .all()
    )
    assert len(rows) == 1
    assert rows[0].status == "PENDING"


def test_upsert_submission_tracking_updates_not_duplicates(db_session, sample_intern, sample_pay_period):
    upsert_submission_tracking(
        db_session, pay_period_id=sample_pay_period.id, intern_id=sample_intern.id, status="PENDING",
    )
    db_session.flush()

    upsert_submission_tracking(
        db_session, pay_period_id=sample_pay_period.id, intern_id=sample_intern.id, status="APPROVED",
    )
    db_session.flush()

    rows = (
        db_session.query(SubmissionTracking)
        .filter_by(pay_period_id=sample_pay_period.id, intern_id=sample_intern.id)
        .all()
    )
    # This assertion is the whole point: exactly one row, not two.
    assert len(rows) == 1
    assert rows[0].status == "APPROVED"


def test_upsert_submission_tracking_rejects_invalid_status(db_session, sample_intern, sample_pay_period):
    with pytest.raises(ValueError):
        upsert_submission_tracking(
            db_session,
            pay_period_id=sample_pay_period.id,
            intern_id=sample_intern.id,
            status="NOT_A_REAL_STATUS",
        )


def test_populate_submission_tracking_is_idempotent(db_session, sample_intern, sample_pay_period):
    first_call = populate_submission_tracking_for_period(
        db_session, sample_pay_period.id, [sample_intern.id]
    )
    second_call = populate_submission_tracking_for_period(
        db_session, sample_pay_period.id, [sample_intern.id]
    )
    assert first_call == 1
    assert second_call == 0
