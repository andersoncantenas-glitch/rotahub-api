CREATE TABLE IF NOT EXISTS plans (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    name TEXT NOT NULL,
    description TEXT,
    monthly_price REAL NOT NULL DEFAULT 0,
    vehicle_limit INTEGER,
    user_limit INTEGER,
    features_json TEXT NOT NULL DEFAULT '{}',
    status TEXT NOT NULL DEFAULT 'active',
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_plans_status
ON plans(status);

CREATE UNIQUE INDEX IF NOT EXISTS idx_plans_code
ON plans(code);
