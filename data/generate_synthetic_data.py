"""
generate_synthetic_data.py

Generates realistic synthetic data for ApprovalFlow development.
Uses Faker for realistic names and emails.
All data is completely fictional — no real institutional data is used.

Run this before db/seed.py to produce the JSON files seed.py loads.

Usage:
    python data/generate_synthetic_data.py
"""

import json
import random
from datetime import date, timedelta
from faker import Faker

fake = Faker()
random.seed(42)
Faker.seed(42)

# --- Configuration ---
NUM_ORGANIZATIONS = 20
NUM_SUPERVISORS = 30       # some orgs have more than one supervisor
NUM_INTERNS = 60           # roughly matches a real pay period cohort
NUM_PAY_PERIODS = 3        # generate a few periods of history


def generate_organizations(n):
    org_types = [
        "Community Health Center", "Public Library", "City Parks Department",
        "Workforce Development Center", "Nonprofit Housing Authority",
        "Community College Foundation", "Regional Transit Authority",
        "County Clerk Office", "Veteran Services Center", "Food Bank",
        "Youth Development Program", "Adult Education Center",
        "Emergency Management Office", "Animal Services Department",
        "Environmental Services Division"
    ]
    orgs = []
    used = set()
    for i in range(n):
        # Mix generated names with realistic org types
        if i < len(org_types):
            name = f"{fake.city()} {org_types[i % len(org_types)]}"
        else:
            name = f"{fake.company()} {random.choice(['Services', 'Institute', 'Center', 'Program'])}"
        # Ensure uniqueness
        while name in used:
            name = f"{fake.company()} Services"
        used.add(name)
        orgs.append({
            "id": i + 1,
            "name": name
        })
    return orgs


def generate_supervisors(n, organizations):
    supervisors = []
    org_ids = [o["id"] for o in organizations]
    used_emails = set()

    for i in range(n):
        first = fake.first_name()
        last = fake.last_name()
        org_id = random.choice(org_ids)
        org = next(o for o in organizations if o["id"] == org_id)

        # Build a realistic work email from org name
        org_domain = org["name"].lower()
        org_domain = "".join(c for c in org_domain if c.isalnum() or c == " ")
        org_domain = org_domain.strip().replace(" ", "")[:12] + ".org"

        base_email = f"{first.lower()}.{last.lower()}@{org_domain}"
        email = base_email
        suffix = 1
        while email in used_emails:
            email = f"{first.lower()}.{last.lower()}{suffix}@{org_domain}"
            suffix += 1
        used_emails.add(email)

        supervisors.append({
            "id": i + 1,
            "first_name": first,
            "last_name": last,
            "email": email,
            "organization_id": org_id,
            "active": True
        })
    return supervisors


def generate_interns(n, supervisors):
    interns = []
    supervisor_ids = [s["id"] for s in supervisors]
    used_emails = set()

    for i in range(n):
        first = fake.first_name()
        last = fake.last_name()
        supervisor_id = random.choice(supervisor_ids)
        supervisor = next(s for s in supervisors if s["id"] == supervisor_id)

        # Institutional email: firstname.lastname@institution.edu
        base_email = f"{first.lower()}.{last.lower()}@institution.edu"
        email = base_email
        suffix = 1
        while email in used_emails:
            email = f"{first.lower()}.{last.lower()}{suffix}@institution.edu"
            suffix += 1
        used_emails.add(email)

        # Secondary student email (used for CC in forwarded approval emails)
        student_id = f"{random.randint(1000, 9999)}"
        student_email = f"{first.lower()[:1]}{last.lower()}{student_id}@student.institution.edu"

        interns.append({
            "id": i + 1,
            "first_name": first,
            "last_name": last,
            "email": email,
            "student_email": student_email,
            "employee_id": f"A{random.randint(10000, 99999)}",
            "supervisor_id": supervisor_id,
            "organization_id": supervisor["organization_id"],
            "start_date": str(fake.date_between(start_date="-1y", end_date="-3m")),
            "active": True
        })
    return interns


def generate_pay_periods(n):
    """
    Generate semi-monthly pay periods (1st-15th and 16th-end of month).
    Submission deadline is the last day of the period.
    Payroll deadline is 2 business days after submission deadline.
    """
    periods = []
    # Start from a few months back
    year = 2026
    month = 3

    for i in range(n):
        if month > 12:
            month = 1
            year += 1

        # First half: 1st to 15th
        start1 = date(year, month, 1)
        end1 = date(year, month, 15)
        payroll1 = end1 + timedelta(days=2)
        periods.append({
            "id": len(periods) + 1,
            "label": f"{start1.strftime('%b')} 1-15 {year}",
            "start_date": str(start1),
            "end_date": str(end1),
            "submission_deadline": str(end1),
            "payroll_deadline": str(payroll1)
        })

        # Second half: 16th to end of month
        import calendar
        last_day = calendar.monthrange(year, month)[1]
        start2 = date(year, month, 16)
        end2 = date(year, month, last_day)
        payroll2 = end2 + timedelta(days=2)
        periods.append({
            "id": len(periods) + 1,
            "label": f"{start2.strftime('%b')} 16-{last_day} {year}",
            "start_date": str(start2),
            "end_date": str(end2),
            "submission_deadline": str(end2),
            "payroll_deadline": str(payroll2)
        })

        month += 1

    return periods[:n * 2]  # n months = n*2 periods


def generate_notification_email(intern, supervisor, pay_period):
    """
    Generate a synthetic Banner-style HTML notification email body.
    This mirrors the actual HTML structure parsed by the VBA MoveItems macro,
    with all real data replaced by synthetic values.

    The real Banner email uses an HTML table where:
    - Column 1: field label
    - Column 6 (index 5): student institutional email (primary lookup key in VBA)
    """
    hours = round(random.uniform(8.0, 40.0), 2)
    period_label = f"{pay_period['start_date']} - {pay_period['end_date']}"

    html = f"""
<html>
<body>
<p>Time Sheet Status as of {pay_period['submission_deadline']}</p>
<table border="1" cellpadding="4" cellspacing="0">
  <tr>
    <td>Employee Name</td>
    <td>{intern['last_name']}, {intern['first_name']}</td>
    <td>Employee ID</td>
    <td>{intern['employee_id']}</td>
    <td>Department</td>
    <td>{intern['email']}</td>
  </tr>
  <tr>
    <td>Pay Period</td>
    <td>{period_label}</td>
    <td>Total Hours</td>
    <td>{hours:.2f}</td>
    <td>Status</td>
    <td>Submitted</td>
  </tr>
</table>
</body>
</html>
""".strip()

    return {
        "subject": f"Time Sheet Notification - {intern['last_name']}, {intern['first_name']}",
        "body_html": html,
        "intern_id": intern["id"],
        "supervisor_id": supervisor["id"],
        "pay_period_id": pay_period["id"],
        "hours": hours
    }


def main():
    print("Generating synthetic data...")

    organizations = generate_organizations(NUM_ORGANIZATIONS)
    print(f"  {len(organizations)} organizations")

    supervisors = generate_supervisors(NUM_SUPERVISORS, organizations)
    print(f"  {len(supervisors)} supervisors")

    interns = generate_interns(NUM_INTERNS, supervisors)
    print(f"  {len(interns)} interns")

    pay_periods = generate_pay_periods(NUM_PAY_PERIODS)
    print(f"  {len(pay_periods)} pay periods")

    # Generate sample notification emails for the most recent pay period
    latest_period = pay_periods[-1]
    active_interns = [i for i in interns if i["active"]]
    # Simulate ~85% submission rate (some interns always miss the deadline)
    submitting_interns = random.sample(active_interns, int(len(active_interns) * 0.85))

    notification_emails = []
    for intern in submitting_interns:
        supervisor = next(s for s in supervisors if s["id"] == intern["supervisor_id"])
        notification_emails.append(
            generate_notification_email(intern, supervisor, latest_period)
        )
    print(f"  {len(notification_emails)} notification emails for pay period: {latest_period['label']}")

    # Write to JSON files for seed.py to load
    import os
    os.makedirs("data/generated", exist_ok=True)

    with open("data/generated/organizations.json", "w") as f:
        json.dump(organizations, f, indent=2)

    with open("data/generated/supervisors.json", "w") as f:
        json.dump(supervisors, f, indent=2)

    with open("data/generated/interns.json", "w") as f:
        json.dump(interns, f, indent=2)

    with open("data/generated/pay_periods.json", "w") as f:
        json.dump(pay_periods, f, indent=2)

    with open("data/generated/notification_emails.json", "w") as f:
        json.dump(notification_emails, f, indent=2)

    print("\nFiles written to data/generated/")
    print("Run 'python db/seed.py' to load into PostgreSQL.")


if __name__ == "__main__":
    main()
