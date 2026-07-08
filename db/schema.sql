-- ApprovalFlow PostgreSQL Schema
-- Replaces Student_Sup_email.xlsx and all its sheets

-- Organizations (replaces employer column in roster)
CREATE TABLE IF NOT EXISTS organizations (
    id          SERIAL PRIMARY KEY,
    name        VARCHAR(255) NOT NULL,
    created_at  TIMESTAMP DEFAULT NOW()
);

-- Supervisors
CREATE TABLE IF NOT EXISTS supervisors (
    id              SERIAL PRIMARY KEY,
    first_name      VARCHAR(100) NOT NULL,
    last_name       VARCHAR(100) NOT NULL,
    email           VARCHAR(255) NOT NULL UNIQUE,
    organization_id INTEGER REFERENCES organizations(id),
    active          BOOLEAN DEFAULT TRUE,
    created_at      TIMESTAMP DEFAULT NOW()
);

-- Interns (replaces Updated_Supervisor email roster sheet)
CREATE TABLE IF NOT EXISTS interns (
    id              SERIAL PRIMARY KEY,
    first_name      VARCHAR(100) NOT NULL,
    last_name       VARCHAR(100) NOT NULL,
    email           VARCHAR(255) NOT NULL UNIQUE,
    student_email   VARCHAR(255),                   -- secondary student address, used for CC
    employee_id     VARCHAR(50),                    -- ID from HR/payroll system
    supervisor_id   INTEGER REFERENCES supervisors(id),
    organization_id INTEGER REFERENCES organizations(id),
    start_date      DATE,
    active          BOOLEAN DEFAULT TRUE,
    created_at      TIMESTAMP DEFAULT NOW()
);

-- Pay periods
CREATE TABLE IF NOT EXISTS pay_periods (
    id              SERIAL PRIMARY KEY,
    label           VARCHAR(50) NOT NULL,           -- e.g. "Jun 1-15 2026"
    start_date      DATE NOT NULL,
    end_date        DATE NOT NULL,
    submission_deadline DATE NOT NULL,
    payroll_deadline    DATE NOT NULL,
    created_at      TIMESTAMP DEFAULT NOW(),
    UNIQUE (start_date, end_date)
);

-- Sent log (replaces Sent_Log sheet)
-- One row per forwarded approval email
CREATE TABLE IF NOT EXISTS sent_log (
    id              SERIAL PRIMARY KEY,
    pay_period_id   INTEGER REFERENCES pay_periods(id),
    intern_id       INTEGER REFERENCES interns(id),
    supervisor_id   INTEGER REFERENCES supervisors(id),
    sent_at         TIMESTAMP NOT NULL,
    email_subject   VARCHAR(500),
    hours_reported  NUMERIC(5,2),
    submission_date DATE,
    created_at      TIMESTAMP DEFAULT NOW()
);

-- Reply log (replaces Reply_Log sheet)
-- One row per parsed supervisor reply
CREATE TABLE IF NOT EXISTS reply_log (
    id              SERIAL PRIMARY KEY,
    sent_log_id     INTEGER REFERENCES sent_log(id),
    pay_period_id   INTEGER REFERENCES pay_periods(id),
    intern_id       INTEGER REFERENCES interns(id),
    supervisor_id   INTEGER REFERENCES supervisors(id),
    received_at     TIMESTAMP NOT NULL,
    response_type   VARCHAR(20) CHECK (response_type IN ('APPROVED', 'REJECTED', 'CORRECTIONS')),
    match_method    VARCHAR(50),                    -- 'cc_match', 'name_score', 'manual'
    reply_body      TEXT,
    created_at      TIMESTAMP DEFAULT NOW()
);

-- Submission tracking (replaces Submission_Tracking sheet)
-- One row per intern per pay period — delete-then-insert, never append
CREATE TABLE IF NOT EXISTS submission_tracking (
    id              SERIAL PRIMARY KEY,
    pay_period_id   INTEGER REFERENCES pay_periods(id),
    intern_id       INTEGER REFERENCES interns(id),
    submitted_at    TIMESTAMP,
    approved_at     TIMESTAMP,
    status          VARCHAR(20) DEFAULT 'PENDING'
                    CHECK (status IN ('PENDING', 'APPROVED', 'REJECTED', 'CORRECTIONS', 'MISSING')),
    notes           TEXT,
    created_at      TIMESTAMP DEFAULT NOW(),
    updated_at      TIMESTAMP DEFAULT NOW(),
    UNIQUE (pay_period_id, intern_id)               -- enforces one row per intern per period
);

-- Index on emails for fast lookups (mirrors LoadLookupDataIntoMemory in VBA)
CREATE INDEX IF NOT EXISTS idx_interns_email ON interns(email);
CREATE INDEX IF NOT EXISTS idx_supervisors_email ON supervisors(email);
CREATE INDEX IF NOT EXISTS idx_sent_log_pay_period ON sent_log(pay_period_id);
CREATE INDEX IF NOT EXISTS idx_submission_tracking_pay_period ON submission_tracking(pay_period_id);
