CREATE TABLE IF NOT EXISTS payments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    subscription_id INTEGER,
    company_id INTEGER NOT NULL,
    amount REAL NOT NULL DEFAULT 0,
    due_date TEXT,
    paid_at TEXT,
    status TEXT NOT NULL DEFAULT 'pending',
    method TEXT,
    reference TEXT,
    boleto_status TEXT,
    boleto_our_number TEXT,
    boleto_digitable_line TEXT,
    boleto_pdf_url TEXT,
    boleto_pdf_path TEXT,
    boleto_generated_at TEXT,
    notes TEXT,
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now')),
    FOREIGN KEY(subscription_id) REFERENCES subscriptions(id),
    FOREIGN KEY(company_id) REFERENCES companies(id)
);

CREATE INDEX IF NOT EXISTS idx_payments_company_status
ON payments(company_id, status);

CREATE INDEX IF NOT EXISTS idx_payments_subscription
ON payments(subscription_id);

CREATE INDEX IF NOT EXISTS idx_payments_due_date
ON payments(due_date);
