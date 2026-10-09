use crate::{drive::mounts::Mount, tasks::TaskPayload};
use std::{
    path::{Path, PathBuf},
    sync::Arc,
    time::Duration,
};

const SESSION_CLOSE_GRACE: Duration = Duration::from_secs(6);

/// Limit session coalescing to formats normally edited by Word, Excel, and
/// PowerPoint. Other files keep the existing one-upload-one-version behavior.
pub(crate) fn is_office_document(path: &Path) -> bool {
    let Some(extension) = path.extension().and_then(|value| value.to_str()) else {
        return false;
    };

    matches!(
        extension.to_ascii_lowercase().as_str(),
        "doc"
            | "docx"
            | "docm"
            | "dot"
            | "dotx"
            | "dotm"
            | "rtf"
            | "xls"
            | "xlsx"
            | "xlsm"
            | "xlsb"
            | "xlt"
            | "xltx"
            | "xltm"
            | "ppt"
            | "pptx"
            | "pptm"
            | "pot"
            | "potx"
            | "potm"
            | "pps"
            | "ppsx"
            | "ppsm"
    )
}

pub(crate) fn is_office_temporary_file(path: &Path) -> bool {
    path.file_name()
        .and_then(|value| value.to_str())
        .is_some_and(|name| name.starts_with("~$") || name.starts_with("~WR"))
}

fn office_document_from_owner_lock(path: &Path) -> Option<PathBuf> {
    let file_name = path.file_name()?.to_str()?;
    let document_name = file_name.strip_prefix("~$")?;
    let document_path = path.with_file_name(document_name);
    is_office_document(&document_path).then_some(document_path)
}

impl Mount {
    /// Build an upload payload that carries the persisted Office session when
    /// one is active. Database failures fall back to the legacy upload path.
    pub(crate) fn upload_payload(&self, path: impl Into<PathBuf>) -> TaskPayload {
        let path = path.into();
        let mut payload = TaskPayload::upload(path.clone());
        if !is_office_document(&path) {
            return payload;
        }

        match self.inventory.version_session_for_path(&self.id, &path) {
            Ok(Some(session_id)) => payload = payload.with_version_session(session_id),
            Ok(None) => {}
            Err(error) => tracing::warn!(
                target: "drive::edit_sessions",
                path = %path.display(),
                error = %error,
                "Failed to load Office version session"
            ),
        }
        payload
    }

    pub(crate) fn office_file_opened(&self, path: &Path) {
        if !is_office_document(path) || is_office_temporary_file(path) {
            return;
        }
        if let Err(error) = self.inventory.edit_session_opened(&self.id, path) {
            tracing::warn!(
                target: "drive::edit_sessions",
                path = %path.display(),
                error = %error,
                "Failed to record Office file open"
            );
        }
    }

    pub(crate) fn office_file_closed(self: &Arc<Self>, path: PathBuf) {
        if !is_office_document(&path) || is_office_temporary_file(&path) {
            return;
        }
        let closed = match self.inventory.edit_session_closed(&self.id, &path) {
            Ok(closed) => closed,
            Err(error) => {
                tracing::warn!(
                    target: "drive::edit_sessions",
                    path = %path.display(),
                    error = %error,
                    "Failed to record Office file close"
                );
                return;
            }
        };
        let Some(closed) = closed else {
            return;
        };

        self.schedule_office_session_finish(path, closed);
    }

    pub(crate) fn office_owner_lock_created(&self, lock_path: &Path) {
        let Some(document_path) = office_document_from_owner_lock(lock_path) else {
            return;
        };
        self.office_file_opened(&document_path);
    }

    pub(crate) fn office_owner_lock_removed(&self, lock_path: &Path) {
        let Some(document_path) = office_document_from_owner_lock(lock_path) else {
            return;
        };
        let closed = match self
            .inventory
            .edit_session_owner_lock_removed(&self.id, &document_path)
        {
            Ok(Some(closed)) => closed,
            Ok(None) => return,
            Err(error) => {
                tracing::warn!(
                    target: "drive::edit_sessions",
                    path = %document_path.display(),
                    error = %error,
                    "Failed to record Office owner lock removal"
                );
                return;
            }
        };
        self.schedule_office_session_finish(document_path, closed);
    }

    fn schedule_office_session_finish(&self, path: PathBuf, session_id: String) {
        let inventory = Arc::clone(&self.inventory);
        let drive_id = self.id.clone();
        tokio::spawn(async move {
            // Filesystem notifications are debounced for two seconds. Retain
            // the session longer so the final Office save uses the same ID.
            tokio::time::sleep(SESSION_CLOSE_GRACE).await;
            if let Err(error) = inventory.finish_edit_session(&drive_id, &path, &session_id) {
                tracing::warn!(
                    target: "drive::edit_sessions",
                    path = %path.display(),
                    error = %error,
                    "Failed to finish Office edit session"
                );
            }
        });
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn detects_supported_office_documents() {
        assert!(is_office_document(Path::new("Report.DOCX")));
        assert!(is_office_document(Path::new("budget.xlsm")));
        assert!(is_office_document(Path::new("slides.pptx")));
        assert!(!is_office_document(Path::new("notes.txt")));
    }

    #[test]
    fn detects_office_lock_files() {
        assert!(is_office_temporary_file(Path::new("~$Report.docx")));
        assert!(!is_office_temporary_file(Path::new("Report.docx")));
    }

    #[test]
    fn maps_owner_lock_to_office_document() {
        assert_eq!(
            office_document_from_owner_lock(Path::new(r"C:\docs\~$Report.docx")),
            Some(PathBuf::from(r"C:\docs\Report.docx"))
        );
        assert_eq!(
            office_document_from_owner_lock(Path::new(r"C:\docs\~$notes.txt")),
            None
        );
    }
}
