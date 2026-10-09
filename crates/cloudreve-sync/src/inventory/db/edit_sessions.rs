use super::InventoryDb;
use anyhow::{Context, Result};
use chrono::Utc;
use diesel::prelude::*;
use std::path::Path;
use uuid::Uuid;

use crate::inventory::schema::edit_sessions::{self, dsl as edit_sessions_dsl};

const STATE_OPEN: &str = "open";
const STATE_CLOSING: &str = "closing";

#[derive(Insertable)]
#[diesel(table_name = edit_sessions)]
struct NewEditSessionRow {
    drive_id: String,
    local_path: String,
    session_id: String,
    open_count: i32,
    state: String,
    updated_at: i64,
}

impl InventoryDb {
    /// Record an Office file handle. Existing rows are deliberately reused so
    /// a session interrupted by a Desktop or Office crash keeps coalescing its
    /// later backups into the same cloud version.
    pub fn edit_session_opened(&self, drive_id: &str, path: &Path) -> Result<String> {
        let mut conn = self.connection()?;
        let local_path = path.to_string_lossy().into_owned();
        let existing = edit_sessions_dsl::edit_sessions
            .filter(edit_sessions_dsl::drive_id.eq(drive_id))
            .filter(edit_sessions_dsl::local_path.eq(&local_path))
            .select((edit_sessions_dsl::session_id, edit_sessions_dsl::open_count))
            .first::<(String, i32)>(&mut conn)
            .optional()
            .context("Failed to query edit session")?;
        let now = Utc::now().timestamp();

        if let Some(existing) = existing {
            let next_count = existing.1.max(0) + 1;
            diesel::update(
                edit_sessions_dsl::edit_sessions
                    .filter(edit_sessions_dsl::drive_id.eq(drive_id))
                    .filter(edit_sessions_dsl::local_path.eq(&local_path)),
            )
            .set((
                edit_sessions_dsl::open_count.eq(next_count),
                edit_sessions_dsl::state.eq(STATE_OPEN),
                edit_sessions_dsl::updated_at.eq(now),
            ))
            .execute(&mut conn)
            .context("Failed to update edit session")?;
            return Ok(existing.0);
        }

        let session_id = Uuid::new_v4().to_string();
        diesel::insert_into(edit_sessions::table)
            .values(NewEditSessionRow {
                drive_id: drive_id.to_string(),
                local_path,
                session_id: session_id.clone(),
                open_count: 1,
                state: STATE_OPEN.to_string(),
                updated_at: now,
            })
            .execute(&mut conn)
            .context("Failed to create edit session")?;
        Ok(session_id)
    }

    /// Record a file close. The row remains in the closing state briefly so a
    /// debounced final filesystem event can still use the same session ID.
    pub fn edit_session_closed(&self, drive_id: &str, path: &Path) -> Result<Option<String>> {
        let mut conn = self.connection()?;
        let local_path = path.to_string_lossy().into_owned();
        let existing = edit_sessions_dsl::edit_sessions
            .filter(edit_sessions_dsl::drive_id.eq(drive_id))
            .filter(edit_sessions_dsl::local_path.eq(&local_path))
            .select((edit_sessions_dsl::session_id, edit_sessions_dsl::open_count))
            .first::<(String, i32)>(&mut conn)
            .optional()
            .context("Failed to query closing edit session")?;
        let Some(existing) = existing else {
            return Ok(None);
        };

        let next_count = (existing.1 - 1).max(0);
        let next_state = if next_count == 0 {
            STATE_CLOSING
        } else {
            STATE_OPEN
        };
        diesel::update(
            edit_sessions_dsl::edit_sessions
                .filter(edit_sessions_dsl::drive_id.eq(drive_id))
                .filter(edit_sessions_dsl::local_path.eq(&local_path)),
        )
        .set((
            edit_sessions_dsl::open_count.eq(next_count),
            edit_sessions_dsl::state.eq(next_state),
            edit_sessions_dsl::updated_at.eq(Utc::now().timestamp()),
        ))
        .execute(&mut conn)
        .context("Failed to close edit session")?;

        Ok((next_count == 0).then_some(existing.0))
    }

    /// Word's Cloud Files open/close callbacks are not always balanced during
    /// atomic saves. Removing its owner lock is the authoritative session end.
    pub fn edit_session_owner_lock_removed(
        &self,
        drive_id: &str,
        path: &Path,
    ) -> Result<Option<String>> {
        let mut conn = self.connection()?;
        let local_path = path.to_string_lossy().into_owned();
        let session_id = edit_sessions_dsl::edit_sessions
            .filter(edit_sessions_dsl::drive_id.eq(drive_id))
            .filter(edit_sessions_dsl::local_path.eq(&local_path))
            .select(edit_sessions_dsl::session_id)
            .first::<String>(&mut conn)
            .optional()
            .context("Failed to query owner-lock edit session")?;
        let Some(session_id) = session_id else {
            return Ok(None);
        };

        diesel::update(
            edit_sessions_dsl::edit_sessions
                .filter(edit_sessions_dsl::drive_id.eq(drive_id))
                .filter(edit_sessions_dsl::local_path.eq(local_path)),
        )
        .set((
            edit_sessions_dsl::open_count.eq(0),
            edit_sessions_dsl::state.eq(STATE_CLOSING),
            edit_sessions_dsl::updated_at.eq(Utc::now().timestamp()),
        ))
        .execute(&mut conn)
        .context("Failed to close owner-lock edit session")?;

        Ok(Some(session_id))
    }

    pub fn version_session_for_path(&self, drive_id: &str, path: &Path) -> Result<Option<String>> {
        let mut conn = self.connection()?;
        let local_path = path.to_string_lossy().into_owned();
        edit_sessions_dsl::edit_sessions
            .filter(edit_sessions_dsl::drive_id.eq(drive_id))
            .filter(edit_sessions_dsl::local_path.eq(local_path))
            .select(edit_sessions_dsl::session_id)
            .first(&mut conn)
            .optional()
            .context("Failed to query version session")
    }

    /// Delete only the closing generation requested by the caller. If the file
    /// was opened again, its state or session ID no longer matches and survives.
    pub fn finish_edit_session(&self, drive_id: &str, path: &Path, session_id: &str) -> Result<()> {
        let mut conn = self.connection()?;
        let local_path = path.to_string_lossy().into_owned();
        diesel::delete(
            edit_sessions_dsl::edit_sessions
                .filter(edit_sessions_dsl::drive_id.eq(drive_id))
                .filter(edit_sessions_dsl::local_path.eq(local_path))
                .filter(edit_sessions_dsl::session_id.eq(session_id))
                .filter(edit_sessions_dsl::state.eq(STATE_CLOSING)),
        )
        .execute(&mut conn)
        .context("Failed to finish edit session")?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    #[test]
    fn session_is_reused_until_last_handle_closes() {
        let dir = tempdir().unwrap();
        let db = InventoryDb::with_path(dir.path().join("inventory.db")).unwrap();
        let path = dir.path().join("report.docx");

        let first = db.edit_session_opened("drive", &path).unwrap();
        let second = db.edit_session_opened("drive", &path).unwrap();
        assert_eq!(first, second);
        assert!(db.edit_session_closed("drive", &path).unwrap().is_none());

        let closed = db
            .edit_session_closed("drive", &path)
            .unwrap()
            .expect("last handle should close the session");
        assert_eq!(closed, first);
        assert_eq!(
            db.version_session_for_path("drive", &path).unwrap(),
            Some(first.clone())
        );

        db.finish_edit_session("drive", &path, &first).unwrap();
        assert_eq!(db.version_session_for_path("drive", &path).unwrap(), None);
    }

    #[test]
    fn interrupted_session_ends_when_database_reopens() {
        let dir = tempdir().unwrap();
        let db_path = dir.path().join("inventory.db");
        let path = dir.path().join("report.xlsx");
        let session = {
            let db = InventoryDb::with_path(db_path.clone()).unwrap();
            db.edit_session_opened("drive", &path).unwrap()
        };

        let reopened = InventoryDb::with_path(db_path).unwrap();
        assert_ne!(
            reopened.edit_session_opened("drive", &path).unwrap(),
            session
        );
    }

    #[test]
    fn owner_lock_removal_closes_unbalanced_handles() {
        let dir = tempdir().unwrap();
        let db = InventoryDb::with_path(dir.path().join("inventory.db")).unwrap();
        let path = dir.path().join("report.docx");
        let session = db.edit_session_opened("drive", &path).unwrap();
        db.edit_session_opened("drive", &path).unwrap();

        let closed = db.edit_session_owner_lock_removed("drive", &path).unwrap();
        assert_eq!(closed, Some(session.clone()));
        db.finish_edit_session("drive", &path, &session).unwrap();
        assert_eq!(db.version_session_for_path("drive", &path).unwrap(), None);
    }
}
