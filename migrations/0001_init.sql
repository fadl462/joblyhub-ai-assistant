PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS organizations (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  row_number INTEGER,
  day_batch TEXT,
  day_number INTEGER,
  batch_position INTEGER,
  planned_send_date TEXT,
  planned_followup_date TEXT,
  name TEXT NOT NULL,
  sector TEXT,
  email TEXT NOT NULL,
  email_normalized TEXT NOT NULL UNIQUE,
  is_ngo INTEGER NOT NULL DEFAULT 0,
  consent_status TEXT NOT NULL DEFAULT 'unknown',
  consent_record TEXT,
  approved INTEGER NOT NULL DEFAULT 0,
  opted_out INTEGER NOT NULL DEFAULT 0,
  unsubscribe_token TEXT NOT NULL UNIQUE,
  email_sent INTEGER NOT NULL DEFAULT 0,
  date_sent TEXT,
  status TEXT NOT NULL DEFAULT 'Pending',
  response_received INTEGER NOT NULL DEFAULT 0,
  response_sentiment TEXT,
  responded_at TEXT,
  followup_sent INTEGER NOT NULL DEFAULT 0,
  followup_sent_at TEXT,
  account_created INTEGER NOT NULL DEFAULT 0,
  job_posted INTEGER NOT NULL DEFAULT 0,
  became_paid INTEGER NOT NULL DEFAULT 0,
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_org_status ON organizations(status);
CREATE INDEX IF NOT EXISTS idx_org_planned_send ON organizations(planned_send_date);
CREATE INDEX IF NOT EXISTS idx_org_followup ON organizations(planned_followup_date);
CREATE INDEX IF NOT EXISTS idx_org_sector ON organizations(sector);

CREATE TABLE IF NOT EXISTS send_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  organization_id INTEGER,
  kind TEXT NOT NULL DEFAULT 'initial',
  mode TEXT NOT NULL DEFAULT 'test',
  sent_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  to_address TEXT,
  success INTEGER NOT NULL DEFAULT 0,
  provider_message_id TEXT,
  error_message TEXT,
  idempotency_key TEXT NOT NULL UNIQUE,
  FOREIGN KEY (organization_id) REFERENCES organizations(id)
);
CREATE INDEX IF NOT EXISTS idx_send_logs_time ON send_logs(sent_at);

CREATE TABLE IF NOT EXISTS settings (
  id INTEGER PRIMARY KEY CHECK (id = 1),
  is_active INTEGER NOT NULL DEFAULT 0,
  test_mode INTEGER NOT NULL DEFAULT 1,
  test_email TEXT NOT NULL DEFAULT 'fadl_rahman3@yahoo.com',
  from_name TEXT NOT NULL DEFAULT 'JoblyHub',
  from_email TEXT NOT NULL DEFAULT 'business@joblyhub.com',
  daily_min INTEGER NOT NULL DEFAULT 10,
  daily_max INTEGER NOT NULL DEFAULT 10,
  window_start_hour INTEGER NOT NULL DEFAULT 8,
  window_end_hour INTEGER NOT NULL DEFAULT 16,
  window_end_minute INTEGER NOT NULL DEFAULT 45,
  followups_enabled INTEGER NOT NULL DEFAULT 1,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
INSERT OR IGNORE INTO settings (id) VALUES (1);

CREATE TABLE IF NOT EXISTS daily_run_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  run_date TEXT NOT NULL UNIQUE,
  scheduled_count INTEGER NOT NULL DEFAULT 0,
  sent_count INTEGER NOT NULL DEFAULT 0,
  completed INTEGER NOT NULL DEFAULT 0,
  started_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  completed_at TEXT
);

CREATE TABLE IF NOT EXISTS notifications (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  title TEXT NOT NULL,
  message TEXT NOT NULL,
  level TEXT NOT NULL DEFAULT 'info',
  is_read INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_notifications_created ON notifications(created_at);
