CREATE TABLE IF NOT EXISTS audit_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    company_id INTEGER,
    user_id INTEGER,
    actor_type TEXT,
    action TEXT NOT NULL,
    entity_type TEXT,
    entity_id TEXT,
    severity TEXT NOT NULL DEFAULT 'info',
    ip_address TEXT,
    metadata_json TEXT NOT NULL DEFAULT '{}',
    created_at TEXT DEFAULT (datetime('now')),
    FOREIGN KEY(company_id) REFERENCES companies(id)
);

CREATE INDEX IF NOT EXISTS idx_audit_logs_company_created
ON audit_logs(company_id, created_at);

CREATE INDEX IF NOT EXISTS idx_audit_logs_action
ON audit_logs(action);
