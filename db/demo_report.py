"""
db/demo_report.py

Prints a readable summary of each pay period in the database, similar in
spirit to the original VBA system's GeneratePayPeriodSummary, but pulling
from the relational schema instead of an Excel workbook.

Usage:
    export DATABASE_URL="postgresql://approvalflow:approvalflow_local@localhost:5432/approvalflow_dev"
    python3 -m db.demo_report
"""

from sqlalchemy import select, func

from db.models import get_session_factory, PayPeriod, Intern, SubmissionTracking


def report_for_period(session, period: PayPeriod) -> None:
    total_interns = session.scalar(select(func.count(Intern.id)))

    status_counts = dict(
        session.execute(
            select(SubmissionTracking.status, func.count(SubmissionTracking.id))
            .where(SubmissionTracking.pay_period_id == period.id)
            .group_by(SubmissionTracking.status)
        ).all()
    )

    submitted = sum(
        v for k, v in status_counts.items() if k != "MISSING"
    )
    approved = status_counts.get("APPROVED", 0)
    rejected = status_counts.get("REJECTED", 0)
    corrections = status_counts.get("CORRECTIONS", 0)
    pending = status_counts.get("PENDING", 0)
    missing = status_counts.get("MISSING", 0)

    print(f"\n{'=' * 50}")
    print(f"PAY PERIOD SUMMARY: {period.label}")
    print(f"{'=' * 50}")
    print(f"{'Total Interns':<25}{total_interns:>10}")
    print(f"{'Submitted':<25}{submitted:>10}")
    print(f"{'Not Submitted':<25}{missing:>10}")
    print(f"{'Approved':<25}{approved:>10}")
    print(f"{'Rejected':<25}{rejected:>10}")
    print(f"{'Corrections Requested':<25}{corrections:>10}")
    print(f"{'Pending Approval':<25}{pending:>10}")

    if submitted > 0:
        response_rate = round((approved + rejected + corrections) / submitted * 100, 1)
        print(f"{'Response Rate':<25}{response_rate:>9}%")


def main():
    Session = get_session_factory()
    session = Session()

    periods = session.scalars(select(PayPeriod).order_by(PayPeriod.start_date)).all()

    if not periods:
        print("No pay periods found. Run 'python3 -m db.seed_synthetic_data' first.")
        return

    for period in periods:
        report_for_period(session, period)

    print()


if __name__ == "__main__":
    main()
