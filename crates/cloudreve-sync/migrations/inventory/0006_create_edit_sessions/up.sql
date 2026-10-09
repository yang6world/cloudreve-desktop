CREATE TABLE IF NOT EXISTS edit_sessions (
    drive_id TEXT NOT NULL,
    local_path TEXT NOT NULL,
    session_id TEXT NOT NULL,
    open_count INTEGER NOT NULL DEFAULT 0,
    state TEXT NOT NULL DEFAULT 'open',
    updated_at BIGINT NOT NULL,
    PRIMARY KEY (drive_id, local_path)
);

CREATE INDEX IF NOT EXISTS idx_edit_sessions_state ON edit_sessions(state);
