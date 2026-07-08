"""
email_parser.py

Parses HR/payroll HTML notification emails to extract intern timesheet data.

This is a NEW implementation for the Python rebuild, not a line-for-line
translation of the VBA system's email parsing. The two differ in real ways:

  - The VBA MoveItems macro (and its ExtractEmailFromHTMLBody helper) counted
    <td> elements across the ENTIRE document, flat, with no table or row
    awareness, and read whichever cell landed at position 6. This parser
    scopes to the first table's first row specifically, a different,
    more structured approach.

  - VBA only ever extracted the student's institutional email from the
    notification HTML. Employee name, ID, pay period, and hours all came
    from a separate Excel roster lookup (LoadLookupDataIntoMemory /
    GetStudentName), not from the notification email itself. This parser
    extracts all of those fields directly from the HTML, which assumes
    the actual HR/payroll notification format includes them. That
    assumption has not yet been checked against a real notification
    email and should be verified before this is relied on.

Uses BeautifulSoup and is stateless — pass in an email body string, get
back a structured dict. This avoids the VBA system's MSHTML DOM parsing,
which caused Outlook to freeze on large batches (see approvalflow-vba's
Known Issues: "Outlook freeze during FastRebuildSentLog").

Test data (tests/test_email_parser.py) is self-contained and does not
reflect a verified real notification format; it exists to test this
parser's own logic, not to validate the assumed HTML structure itself.

"""

from bs4 import BeautifulSoup
from dataclasses import dataclass
from datetime import date
from typing import Optional
import re


@dataclass
class ParsedNotification:
    """Structured result of parsing a Banner notification email."""
    employee_name: Optional[str]       # "Last, First" format from Banner
    first_name: Optional[str]
    last_name: Optional[str]
    employee_id: Optional[str]         # HR/payroll system ID
    institutional_email: Optional[str] # Primary lookup key — TD index 5 in VBA
    pay_period_start: Optional[str]
    pay_period_end: Optional[str]
    total_hours: Optional[float]
    submission_date: Optional[date]    # Parsed from "Time Sheet Status as of [date]"
    raw_html: str
    parse_errors: list


def parse_notification_email(html_body: str) -> ParsedNotification:
    """
    Parse a Banner-style HTML notification email body.

    Args:
        html_body: Raw HTML string from the notification email body

    Returns:
        ParsedNotification dataclass with extracted fields.
        Fields are None if extraction failed; parse_errors lists what went wrong.

    VBA equivalent: the HTML parsing block inside MoveItems
    """
    errors = []
    soup = BeautifulSoup(html_body, "lxml")

    # --- Extract submission date from status line ---
    # Banner includes a line: "Time Sheet Status as of MM/DD/YYYY"
    # VBA equivalent: ExtractTimesheetStatusDate function
    # Critical lesson: this is the SUBMISSION date, not the pay period date.
    # A timesheet submitted on May 5 for the April 16-30 period has a status
    # date of May 5 -- do not use this as a pay period filter.
    submission_date = None
    status_text = soup.get_text()
    status_match = re.search(
        r"Time Sheet Status as of (\d{1,2}/\d{1,2}/\d{4})",
        status_text,
        re.IGNORECASE
    )
    if status_match:
        try:
            from datetime import datetime
            submission_date = datetime.strptime(
                status_match.group(1), "%m/%d/%Y"
            ).date()
        except ValueError as e:
            errors.append(f"Could not parse submission date: {e}")
    else:
        errors.append("Submission date not found (expected 'Time Sheet Status as of MM/DD/YYYY')")

    # --- Extract fields from HTML table ---
    # Banner notification emails use a two-row HTML table.
    # Row 1: Employee Name | value | Employee ID | value | Department | student_email
    # Row 2: Pay Period    | value | Total Hours | value | Status     | value
    #
    # VBA used flat, document-wide TD indexing (Item(5)) for this same field; this implementation scopes to the first table's first row instead.
    employee_name = None
    employee_id = None
    institutional_email = None
    pay_period_raw = None
    total_hours = None

    tables = soup.find_all("table")
    if not tables:
        errors.append("No HTML table found in email body")
        return ParsedNotification(
            employee_name=None, first_name=None, last_name=None,
            employee_id=None, institutional_email=None,
            pay_period_start=None, pay_period_end=None,
            total_hours=None, submission_date=submission_date,
            raw_html=html_body, parse_errors=errors
        )

    # Use first table
    table = tables[0]
    rows = table.find_all("tr")

    # Row 1: name, ID, email
    if len(rows) >= 1:
        cells = rows[0].find_all("td")
        cell_texts = [c.get_text(strip=True) for c in cells]

        # Cell index 1: employee name (label is index 0)
        if len(cell_texts) > 1:
            employee_name = cell_texts[1] if cell_texts[1] else None

        # Cell index 3: employee ID (label is index 2)
        if len(cell_texts) > 3:
            employee_id = cell_texts[3] if cell_texts[3] else None

        # Cell index 5: institutional email — primary lookup key
        # VBA used flat, document-wide TD indexing (Item(5)) for this same field, no table/row awareness; this parser scopes to the first table's first row instead
        if len(cell_texts) > 5:
            raw_email = cell_texts[5]
            if "@" in raw_email:
                institutional_email = raw_email.lower().strip()
            else:
                errors.append(
                    f"Cell index 5 does not look like an email address: '{raw_email}'"
                )
        else:
            errors.append(
                f"Row 1 has only {len(cell_texts)} cells; expected at least 6 for email at index 5"
            )

    # Row 2: pay period, hours
    if len(rows) >= 2:
        cells = rows[1].find_all("td")
        cell_texts = [c.get_text(strip=True) for c in cells]

        # Cell index 1: pay period string "YYYY-MM-DD - YYYY-MM-DD"
        if len(cell_texts) > 1:
            pay_period_raw = cell_texts[1]

        # Cell index 3: total hours
        if len(cell_texts) > 3:
            try:
                total_hours = float(cell_texts[3])
            except ValueError:
                errors.append(f"Could not parse hours from: '{cell_texts[3]}'")

    # --- Parse name into first/last ---
    first_name = None
    last_name = None
    if employee_name:
        # Banner format: "Last, First" — split on first comma
        if "," in employee_name:
            parts = employee_name.split(",", 1)
            last_name = parts[0].strip()
            first_name = parts[1].strip()
        else:
            errors.append(
                f"Employee name not in 'Last, First' format: '{employee_name}'"
            )
            # Fall back: treat entire string as last name
            last_name = employee_name.strip()

    # --- Parse pay period dates ---
    pay_period_start = None
    pay_period_end = None
    if pay_period_raw:
        # Expected format: "YYYY-MM-DD - YYYY-MM-DD"
        date_match = re.search(
            r"(\d{4}-\d{2}-\d{2})\s*[-–]\s*(\d{4}-\d{2}-\d{2})",
            pay_period_raw
        )
        if date_match:
            pay_period_start = date_match.group(1)
            pay_period_end = date_match.group(2)
        else:
            errors.append(
                f"Could not parse pay period dates from: '{pay_period_raw}'"
            )

    return ParsedNotification(
        employee_name=employee_name,
        first_name=first_name,
        last_name=last_name,
        employee_id=employee_id,
        institutional_email=institutional_email,
        pay_period_start=pay_period_start,
        pay_period_end=pay_period_end,
        total_hours=total_hours,
        submission_date=submission_date,
        raw_html=html_body,
        parse_errors=errors
    )


def strip_quoted_content(email_body: str) -> str:
    """
    Remove quoted original message content from a reply email body.

    Critical lesson from VBA: DetermineResponseType misclassified APPROVE
    replies as REJECTED because the quoted original message contained the word
    "REJECTED" in the instructions. Always strip quoted content before
    keyword classification.

    Strips at common reply separators:
    - "From:" at start of line
    - A line of underscores (______)
    - "-----Original Message-----"
    - "Sent:" at start of line
    - "On [date/time]" patterns
    """
    separators = [
        r"^From:\s",
        r"^_{5,}",
        r"^-{5,}Original Message-{5,}",
        r"^Sent:\s",
        r"^On .+ wrote:",
    ]
    pattern = re.compile(
        "|".join(separators),
        re.IGNORECASE | re.MULTILINE
    )
    match = pattern.search(email_body)
    if match:
        return email_body[:match.start()].strip()
    return email_body.strip()


if __name__ == "__main__":
    # Quick smoke test against a synthetic email
    sample_html = """
    <html><body>
    <p>Time Sheet Status as of 06/15/2026</p>
    <table border="1">
      <tr>
        <td>Employee Name</td>
        <td>Doe, Jane</td>
        <td>Employee ID</td>
        <td>A12345</td>
        <td>Department</td>
        <td>jane.doe@institution.edu</td>
      </tr>
      <tr>
        <td>Pay Period</td>
        <td>2026-06-01 - 2026-06-15</td>
        <td>Total Hours</td>
        <td>32.50</td>
        <td>Status</td>
        <td>Submitted</td>
      </tr>
    </table>
    </body></html>
    """

    result = parse_notification_email(sample_html)
    print("Parse result:")
    print(f"  Name:        {result.first_name} {result.last_name}")
    print(f"  Email:       {result.institutional_email}")
    print(f"  Employee ID: {result.employee_id}")
    print(f"  Hours:       {result.total_hours}")
    print(f"  Period:      {result.pay_period_start} to {result.pay_period_end}")
    print(f"  Submitted:   {result.submission_date}")
    print(f"  Errors:      {result.parse_errors}")

    # Test quoted content stripping
    reply_with_quote = """APPROVE

Looks good.

From: Payroll <payroll@institution.edu>
To: supervisor@org.org
Subject: Timesheet Approval Request

Please APPROVE or REJECT the following timesheet...
    """
    stripped = strip_quoted_content(reply_with_quote)
    print(f"\nStripped reply body:\n{stripped}")
