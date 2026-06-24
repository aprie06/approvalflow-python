"""
test_email_parser.py

Tests for core/email_parser.py

Run with: pytest tests/test_email_parser.py -v
"""

import pytest
from core.email_parser import parse_notification_email, strip_quoted_content


# --- Fixtures ---

VALID_EMAIL_HTML = """
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

MISSING_EMAIL_HTML = """
<html><body>
<p>Time Sheet Status as of 06/15/2026</p>
<table border="1">
  <tr>
    <td>Employee Name</td>
    <td>Doe, Jane</td>
    <td>Employee ID</td>
    <td>A12345</td>
    <td>Department</td>
    <td>NOT_AN_EMAIL</td>
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

NO_TABLE_HTML = """
<html><body>
<p>Time Sheet Status as of 06/15/2026</p>
<p>No table here.</p>
</body></html>
"""


# --- Tests: parse_notification_email ---

class TestParseNotificationEmail:

    def test_parses_name_correctly(self):
        result = parse_notification_email(VALID_EMAIL_HTML)
        assert result.first_name == "Jane"
        assert result.last_name == "Doe"

    def test_parses_institutional_email(self):
        result = parse_notification_email(VALID_EMAIL_HTML)
        assert result.institutional_email == "jane.doe@institution.edu"

    def test_parses_employee_id(self):
        result = parse_notification_email(VALID_EMAIL_HTML)
        assert result.employee_id == "A12345"

    def test_parses_hours(self):
        result = parse_notification_email(VALID_EMAIL_HTML)
        assert result.total_hours == 32.50

    def test_parses_pay_period_dates(self):
        result = parse_notification_email(VALID_EMAIL_HTML)
        assert result.pay_period_start == "2026-06-01"
        assert result.pay_period_end == "2026-06-15"

    def test_parses_submission_date(self):
        result = parse_notification_email(VALID_EMAIL_HTML)
        from datetime import date
        assert result.submission_date == date(2026, 6, 15)

    def test_no_errors_on_valid_email(self):
        result = parse_notification_email(VALID_EMAIL_HTML)
        assert result.parse_errors == []

    def test_records_error_when_email_missing(self):
        result = parse_notification_email(MISSING_EMAIL_HTML)
        assert result.institutional_email is None
        assert any("Cell index 5" in e for e in result.parse_errors)

    def test_records_error_when_no_table(self):
        result = parse_notification_email(NO_TABLE_HTML)
        assert any("No HTML table" in e for e in result.parse_errors)

    def test_email_lowercased(self):
        html = VALID_EMAIL_HTML.replace(
            "jane.doe@institution.edu", "Jane.Doe@Institution.EDU"
        )
        result = parse_notification_email(html)
        assert result.institutional_email == "jane.doe@institution.edu"


# --- Tests: strip_quoted_content ---

class TestStripQuotedContent:

    def test_strips_from_from_line(self):
        body = "APPROVE\n\nFrom: payroll@institution.edu\nOriginal message here"
        result = strip_quoted_content(body)
        assert result == "APPROVE"

    def test_strips_from_underscores(self):
        body = "APPROVE\n\n______________________________\nOriginal message here"
        result = strip_quoted_content(body)
        assert result == "APPROVE"

    def test_strips_from_original_message_separator(self):
        body = "APPROVE\n\n-----Original Message-----\nOriginal message here"
        result = strip_quoted_content(body)
        assert result == "APPROVE"

    def test_no_quoted_content_unchanged(self):
        body = "APPROVE\n\nHours look correct."
        result = strip_quoted_content(body)
        assert result == "APPROVE\n\nHours look correct."

    def test_approve_not_contaminated_by_reject_in_quote(self):
        """
        Critical regression test.
        VBA bug: DetermineResponseType classified APPROVE replies as REJECTED
        because the quoted original message contained the word REJECTED in
        the instructions. Stripping quoted content before classification
        prevents this.
        """
        body = (
            "APPROVE\n\n"
            "From: payroll@institution.edu\n"
            "Please reply with APPROVE or REJECT.\n"
            "If REJECTED, include a reason."
        )
        stripped = strip_quoted_content(body)
        assert "REJECT" not in stripped
        assert "APPROVE" in stripped
