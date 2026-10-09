use axum::{
    extract::{Query, State},
    http::StatusCode,
    response::IntoResponse,
    routing::{get, post},
    Json, Router,
};
use chrono::{DateTime, Utc};
use cloudreve_sync::{
    uploader::{ProgressCallback, ProgressUpdate, UploadParams, Uploader, UploaderConfig},
    DriveManager,
};
use cloudreve_api::{
    api::ExplorerApi,
    models::{
        explorer::{file_type, FileResponse, FileURLService, GetFileInfoService, ListFileService, VersionControlService},
        uri::CrUri,
    },
};
use serde::{Deserialize, Serialize};
use std::{
    collections::HashMap,
    fs,
    path::{Path, PathBuf},
    sync::Arc,
    time::{SystemTime, UNIX_EPOCH},
};
use tokio::{sync::Mutex, time::{sleep, Duration}};
use uuid::Uuid;

const OFFICE_BRIDGE_PORT: u16 = 5213;
const OFFICE_SYNC_FILE: &str = "office-documents.json";
const OFFICE_SYNC_INTERVAL: Duration = Duration::from_secs(2);
/// Grace period a restore allows for Office to release a document it just closed.
const OFFICE_RESTORE_SESSION_WAIT: Duration = Duration::from_secs(25);
const OFFICE_RESTORE_SESSION_POLL: Duration = Duration::from_millis(400);

#[derive(Clone)]
struct BridgeState {
    drive_manager: Arc<DriveManager>,
    documents: Arc<Mutex<OfficeSyncRegistry>>,
    sync_locks: Arc<Mutex<HashMap<String, Arc<Mutex<()>>>>>,
    registry_path: Arc<PathBuf>,
}

#[derive(Debug, Default, Serialize, Deserialize)]
struct OfficeSyncRegistry {
    #[serde(default)]
    documents: HashMap<String, OfficeManagedDocument>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct OfficeManagedDocument {
    local_path: String,
    drive_id: String,
    folder_uri: String,
    remote_uri: String,
    #[serde(default)]
    active_session: Option<String>,
    #[serde(default)]
    last_uploaded_modified: Option<i64>,
    #[serde(default)]
    last_remote_entity: Option<String>,
    #[serde(default)]
    last_synced_at: Option<String>,
    #[serde(default)]
    last_error: Option<String>,
}

#[derive(Debug, Deserialize)]
pub struct OfficePathQuery {
    path: String,
}

#[derive(Debug, Deserialize)]
pub struct RestoreRequest {
    path: String,
    version: String,
}

#[derive(Debug, Deserialize)]
struct FolderQuery {
    drive_id: String,
    path: Option<String>,
}

#[derive(Debug, Deserialize)]
struct StartSyncRequest {
    path: String,
    drive_id: String,
    folder_uri: String,
}

#[derive(Debug, Deserialize)]
struct StopSyncRequest {
    path: String,
}

#[derive(Debug, Serialize)]
pub struct OfficeVersion {
    pub id: String,
    pub size: i64,
    pub created_at: String,
    pub current: bool,
}

#[derive(Debug, Serialize)]
pub struct OfficeFileResponse {
    pub name: String,
    pub path: String,
    pub size: i64,
    pub updated_at: String,
    pub current_version: Option<String>,
    pub versions: Vec<OfficeVersion>,
    pub managed: bool,
    pub external: bool,
    pub remote_uri: Option<String>,
    pub last_synced_at: Option<String>,
    pub last_error: Option<String>,
    pub document_open: bool,
}

#[derive(Debug, Serialize)]
struct OfficeDriveResponse {
    id: String,
    name: String,
    instance_url: String,
    root_uri: String,
}

#[derive(Debug, Serialize)]
struct OfficeFolderResponse {
    name: String,
    path: String,
}

#[derive(Debug, Serialize)]
struct OfficeFolderListResponse {
    current_path: String,
    parent_path: Option<String>,
    folders: Vec<OfficeFolderResponse>,
}

/// Upload progress is intentionally internal; Office receives only completed state via polling.
struct OfficeUploadProgress;

impl ProgressCallback for OfficeUploadProgress {
    fn on_progress(&self, _update: ProgressUpdate) {}
}

pub async fn start(drive_manager: Arc<DriveManager>) {
    let registry_path = match office_registry_path() {
        Ok(path) => path,
        Err(error) => {
            tracing::warn!(target: "office_bridge", %error, "Failed to initialize Office document registry");
            return;
        }
    };
    let state = BridgeState {
        drive_manager,
        documents: Arc::new(Mutex::new(load_registry(&registry_path))),
        sync_locks: Arc::new(Mutex::new(HashMap::new())),
        registry_path: Arc::new(registry_path),
    };
    tokio::spawn(run_sync_worker(state.clone()));

    let router = Router::new()
        .route("/api/v1/office/health", get(health))
        .route("/api/v1/office/versions", get(versions))
        .route("/api/v1/office/restore", post(restore))
        .route("/api/v1/office/drives", get(drives))
        .route("/api/v1/office/folders", get(folders))
        .route("/api/v1/office/sync", post(start_sync))
        .route("/api/v1/office/sync/disable", post(stop_sync))
        .with_state(state);
    let listener = match tokio::net::TcpListener::bind(("127.0.0.1", OFFICE_BRIDGE_PORT)).await {
        Ok(listener) => listener,
        Err(error) => {
            tracing::warn!(target: "office_bridge", %error, port = OFFICE_BRIDGE_PORT, "Failed to bind Office add-in bridge");
            return;
        }
    };
    tracing::info!(target: "office_bridge", port = OFFICE_BRIDGE_PORT, "Office add-in bridge listening");
    if let Err(error) = axum::serve(listener, router).await {
        tracing::warn!(target: "office_bridge", %error, "Office add-in bridge stopped");
    }
}

async fn health() -> impl IntoResponse {
    Json(serde_json::json!({"service": "cloudreve-desktop", "office_bridge": true}))
}

async fn versions(
    State(state): State<BridgeState>,
    Query(query): Query<OfficePathQuery>,
) -> Result<Json<OfficeFileResponse>, BridgeError> {
    let path = PathBuf::from(&query.path);
    let document = managed_document(&state, &path).await;
    let remote_file = match &document {
        Some(document) => get_remote_file(&state, document).await?,
        None => match state.drive_manager.office_file_info(path.clone()).await {
            Ok(file) => Some(file),
            Err(error) if error.to_string().contains("No Cloudreve drive") => None,
            Err(error) => return Err(error.into()),
        },
    };
    Ok(Json(build_file_response(path, remote_file, document)))
}

async fn restore(
    State(state): State<BridgeState>,
    Json(request): Json<RestoreRequest>,
) -> Result<StatusCode, BridgeError> {
    let path = PathBuf::from(request.path);
    if let Some(document) = managed_document(&state, &path).await {
        restore_managed_document(&state, &document, &request.version).await?;
    } else {
        state.drive_manager.office_restore_version(path, request.version).await?;
    }
    Ok(StatusCode::NO_CONTENT)
}

async fn drives(State(state): State<BridgeState>) -> Result<Json<Vec<OfficeDriveResponse>>, BridgeError> {
    let mut result = Vec::new();
    for drive in state.drive_manager.list_drives().await {
        let root_uri = CrUri::new(&drive.remote_path)
            .map_err(|error| anyhow::anyhow!("Invalid Cloudreve drive path: {error}"))?
            .base(true);
        result.push(OfficeDriveResponse {
            id: drive.id,
            name: drive.name,
            instance_url: drive.instance_url,
            root_uri,
        });
    }
    Ok(Json(result))
}

async fn folders(
    State(state): State<BridgeState>,
    Query(query): Query<FolderQuery>,
) -> Result<Json<OfficeFolderListResponse>, BridgeError> {
    let mount = state
        .drive_manager
        .get_drive(&query.drive_id)
        .await
        .ok_or_else(|| anyhow::anyhow!("Cloudreve account is no longer available"))?;
    let config = mount.get_config().await;
    let root_uri = CrUri::new(&config.remote_path)?.base(true);
    let current_path = query.path.unwrap_or_else(|| root_uri.clone());
    if !is_within_root(&current_path, &root_uri) {
        return Err(anyhow::anyhow!("The selected folder is outside the current Cloudreve account").into());
    }
    let response = mount
        .cr_client
        .list_files(&ListFileService {
            uri: current_path.clone(),
            page: Some(1),
            page_size: Some(200),
            order_by: Some("name".to_string()),
            order_direction: Some("asc".to_string()),
            next_page_token: None,
        })
        .await
        .map_err(|error| anyhow::anyhow!(error.to_string()))?;
    let folders = response
        .files
        .into_iter()
        .filter(|file| file.file_type == file_type::FOLDER)
        .map(|file| OfficeFolderResponse { name: file.name, path: file.path })
        .collect();
    let parent_path = CrUri::new(&current_path)
        .ok()
        .and_then(|uri| uri.parent().ok())
        .map(|uri| uri.to_string())
        .filter(|parent| is_within_root(parent, &root_uri));
    Ok(Json(OfficeFolderListResponse { current_path, parent_path, folders }))
}

async fn start_sync(
    State(state): State<BridgeState>,
    Json(request): Json<StartSyncRequest>,
) -> Result<Json<OfficeFileResponse>, BridgeError> {
    let path = PathBuf::from(&request.path);
    let file_name = path
        .file_name()
        .and_then(|name| name.to_str())
        .filter(|name| !name.is_empty())
        .ok_or_else(|| anyhow::anyhow!("The Office document must be saved before enabling sync"))?;
    if !path.is_file() {
        return Err(anyhow::anyhow!("The Office document must be saved before enabling sync").into());
    }
    let mount = state
        .drive_manager
        .get_drive(&request.drive_id)
        .await
        .ok_or_else(|| anyhow::anyhow!("Cloudreve account is no longer available"))?;
    let config = mount.get_config().await;
    let root_uri = CrUri::new(&config.remote_path)?.base(true);
    if !is_within_root(&request.folder_uri, &root_uri) {
        return Err(anyhow::anyhow!("The selected folder is outside the current Cloudreve account").into());
    }
    let mut remote_uri = CrUri::new(&request.folder_uri)?;
    remote_uri.join(&[file_name]);
    let key = path_key(&path);
    {
        let mut registry = state.documents.lock().await;
        registry.documents.insert(
            key,
            OfficeManagedDocument {
                local_path: path.to_string_lossy().into_owned(),
                drive_id: request.drive_id,
                folder_uri: request.folder_uri,
                remote_uri: remote_uri.to_string(),
                active_session: Some(Uuid::new_v4().to_string()),
                last_uploaded_modified: None,
                last_remote_entity: None,
                last_synced_at: None,
                last_error: None,
            },
        );
        save_registry(&state.registry_path, &registry)?;
    }
    sync_document(&state, &path).await?;
    let document = managed_document(&state, &path)
        .await
        .ok_or_else(|| anyhow::anyhow!("Office document registration was lost"))?;
    let remote_file = get_remote_file(&state, &document).await?;
    Ok(Json(build_file_response(path, remote_file, Some(document))))
}

async fn stop_sync(
    State(state): State<BridgeState>,
    Json(request): Json<StopSyncRequest>,
) -> Result<StatusCode, BridgeError> {
    let path = PathBuf::from(request.path);
    let mut registry = state.documents.lock().await;
    if registry.documents.remove(&path_key(&path)).is_none() {
        return Err(anyhow::anyhow!("This document is not managed by Cloudreve").into());
    }
    save_registry(&state.registry_path, &registry)?;
    Ok(StatusCode::NO_CONTENT)
}

async fn run_sync_worker(state: BridgeState) {
    loop {
        let paths = {
            let registry = state.documents.lock().await;
            registry
                .documents
                .values()
                .map(|document| PathBuf::from(&document.local_path))
                .collect::<Vec<_>>()
        };
        for path in paths {
            if let Err(error) = sync_document(&state, &path).await {
                tracing::warn!(target: "office_bridge", path = %path.display(), %error, "Office document backup failed");
                continue;
            }
            if let Err(error) = sync_remote_document(&state, &path).await {
                tracing::warn!(target: "office_bridge", path = %path.display(), %error, "Office document download failed");
            }
        }
        sleep(OFFICE_SYNC_INTERVAL).await;
    }
}

/// Backs up a registered document only after its on-disk content changes.
async fn sync_document(state: &BridgeState, path: &Path) -> anyhow::Result<()> {
    // The worker and an explicit enable-sync request can reach the same file at
    // once. Serialize each document so one disk change creates one upload.
    let document_lock = {
        let mut locks = state.sync_locks.lock().await;
        locks
            .entry(path_key(path))
            .or_insert_with(|| Arc::new(Mutex::new(())))
            .clone()
    };
    let _document_guard = document_lock.lock().await;
    let metadata = match fs::metadata(path) {
        Ok(metadata) if metadata.is_file() => metadata,
        Ok(_) => anyhow::bail!("Registered Office path is not a file"),
        Err(error) => return update_document_error(state, path, error.to_string()).await,
    };
    let modified = system_time_millis(metadata.modified().unwrap_or(UNIX_EPOCH));
    let lock_present = office_lock_path(path).is_some_and(|lock| lock.exists());
    let (document, needs_upload, should_close_session) = {
        let mut registry = state.documents.lock().await;
        let document = registry
            .documents
            .get_mut(&path_key(path))
            .ok_or_else(|| anyhow::anyhow!("Office document is not registered"))?;
        if lock_present && document.active_session.is_none() {
            document.active_session = Some(Uuid::new_v4().to_string());
        }
        let changed = document.last_uploaded_modified != Some(modified);
        let should_close_session = !lock_present && document.active_session.is_some();
        (document.clone(), changed, should_close_session)
    };
    if !needs_upload {
        if should_close_session {
            finish_document_session(state, path, false).await?;
        }
        return Ok(());
    }

    let mount = state
        .drive_manager
        .get_drive(&document.drive_id)
        .await
        .ok_or_else(|| anyhow::anyhow!("Cloudreve account is no longer available"))?;
    let existing = mount
        .cr_client
        .get_file_info(&GetFileInfoService {
            uri: Some(document.remote_uri.clone()),
            id: None,
            extended: None,
            folder_summary: None,
        })
        .await
        .ok();
    let uploader = Uploader::new(mount.cr_client.clone(), mount.inventory.clone(), UploaderConfig::default());
    let previous_version = existing
        .as_ref()
        .and_then(|file| file.primary_entity.clone())
        .unwrap_or_default();
    uploader
        .upload(
            UploadParams {
                local_path: path.to_path_buf(),
                remote_uri: document.remote_uri.clone(),
                file_size: metadata.len(),
                mime_type: None,
                last_modified: Some(modified),
                overwrite: existing.is_some(),
                previous_version,
                version_session: document.active_session.clone(),
                task_id: format!("office-{}", Uuid::new_v4()),
                drive_id: document.drive_id.clone(),
            },
            OfficeUploadProgress,
        )
        .await?;

    let current_entity = mount
        .cr_client
        .get_file_info(&GetFileInfoService {
            uri: Some(document.remote_uri.clone()),
            id: None,
            extended: None,
            folder_summary: None,
        })
        .await
        .ok()
        .and_then(|file| file.primary_entity);
    let mut registry = state.documents.lock().await;
    if let Some(document) = registry.documents.get_mut(&path_key(path)) {
        document.last_uploaded_modified = Some(modified);
        document.last_remote_entity = current_entity;
        document.last_synced_at = Some(Utc::now().to_rfc3339());
        document.last_error = None;
        if !lock_present {
            document.active_session = None;
        }
        save_registry(&state.registry_path, &registry)?;
    }
    Ok(())
}

/// Pull a newer cloud version only while the local document is closed and clean.
async fn sync_remote_document(state: &BridgeState, path: &Path) -> anyhow::Result<()> {
    if office_lock_path(path).is_some_and(|lock| lock.exists()) {
        return Ok(());
    }
    let metadata = fs::metadata(path)?;
    let local_modified = system_time_millis(metadata.modified().unwrap_or(UNIX_EPOCH));
    let document = managed_document(state, path)
        .await
        .ok_or_else(|| anyhow::anyhow!("Office document is not registered"))?;
    if document.active_session.is_some() || document.last_uploaded_modified != Some(local_modified) {
        return Ok(());
    }
    let mount = state
        .drive_manager
        .get_drive(&document.drive_id)
        .await
        .ok_or_else(|| anyhow::anyhow!("Cloudreve account is no longer available"))?;
    let remote = mount
        .cr_client
        .get_file_info(&GetFileInfoService {
            uri: Some(document.remote_uri.clone()),
            id: None,
            extended: None,
            folder_summary: None,
        })
        .await?;
    let Some(remote_entity) = remote.primary_entity else {
        return Ok(());
    };
    if document.last_remote_entity.as_deref() == Some(remote_entity.as_str()) {
        return Ok(());
    }
    download_document_content(&mount, &document.remote_uri, path).await?;
    let modified = fs::metadata(path)
        .ok()
        .and_then(|metadata| metadata.modified().ok())
        .map(system_time_millis);
    let mut registry = state.documents.lock().await;
    if let Some(current) = registry.documents.get_mut(&path_key(path)) {
        current.last_uploaded_modified = modified;
        current.last_remote_entity = Some(remote_entity);
        current.last_synced_at = Some(Utc::now().to_rfc3339());
        current.last_error = None;
        save_registry(&state.registry_path, &registry)?;
    }
    Ok(())
}

async fn restore_managed_document(
    state: &BridgeState,
    document: &OfficeManagedDocument,
    version: &str,
) -> anyhow::Result<()> {
    // The add-in closes the document right before restoring. The worker only
    // clears the editing session on its next pass, and the final save of that
    // session still has to land, so give both a moment to complete instead of
    // rejecting a restore that is about to become valid.
    let document = wait_for_closed_document(state, document).await?;
    let mount = state
        .drive_manager
        .get_drive(&document.drive_id)
        .await
        .ok_or_else(|| anyhow::anyhow!("Cloudreve account is no longer available"))?;
    mount
        .cr_client
        .set_current_version(&VersionControlService { uri: document.remote_uri.clone(), version: version.to_string() })
        .await?;
    download_document_content(&mount, &document.remote_uri, &PathBuf::from(&document.local_path)).await?;
    let remote = mount
        .cr_client
        .get_file_info(&GetFileInfoService {
            uri: Some(document.remote_uri.clone()),
            id: None,
            extended: None,
            folder_summary: None,
        })
        .await?;
    let path = PathBuf::from(&document.local_path);
    let modified = fs::metadata(&path)
        .ok()
        .and_then(|metadata| metadata.modified().ok())
        .map(system_time_millis);
    let mut registry = state.documents.lock().await;
    if let Some(current) = registry.documents.get_mut(&path_key(&path)) {
        current.last_uploaded_modified = modified;
        current.last_remote_entity = remote.primary_entity;
        current.last_synced_at = Some(Utc::now().to_rfc3339());
        current.last_error = None;
        save_registry(&state.registry_path, &registry)?;
    }
    Ok(())
}

/// Wait for the document to be closed and its pending backup to finish, then
/// return its up-to-date registration.
async fn wait_for_closed_document(
    state: &BridgeState,
    document: &OfficeManagedDocument,
) -> anyhow::Result<OfficeManagedDocument> {
    let path = PathBuf::from(&document.local_path);
    let deadline = tokio::time::Instant::now() + OFFICE_RESTORE_SESSION_WAIT;
    loop {
        let lock_present = office_lock_path(&path).is_some_and(|lock| lock.exists());
        let current = managed_document(state, &path)
            .await
            .ok_or_else(|| anyhow::anyhow!("This document is not managed by Cloudreve"))?;
        if !lock_present && current.active_session.is_none() {
            return Ok(current);
        }
        if tokio::time::Instant::now() >= deadline {
            anyhow::bail!("The Office document is still open; close it before restoring a version");
        }
        if !lock_present {
            // Flush the final save of the closed session so it cannot land on
            // top of the version selected below.
            let _ = sync_document(state, &path).await;
        }
        sleep(OFFICE_RESTORE_SESSION_POLL).await;
    }
}

/// Downloads the selected current entity to a sibling temporary file before replacing the document.
async fn download_document_content(
    mount: &cloudreve_sync::drive::mounts::Mount,
    remote_uri: &str,
    path: &Path,
) -> anyhow::Result<()> {
    let links = mount
        .cr_client
        .get_file_url(&FileURLService {
            uris: vec![remote_uri.to_string()],
            download: Some(true),
            // Cloudreve returns an empty redirect response when this is true.
            // Request the JSON URL payload, then download through reqwest.
            redirect: Some(false),
            entity: None,
            no_cache: Some(true),
            skip_error: None,
            use_primary_site_url: None,
            archive: None,
        })
        .await?;
    let url = links.urls.first().ok_or_else(|| anyhow::anyhow!("Cloudreve did not return a download URL"))?.url.clone();
    let content = reqwest::get(url).await?.error_for_status()?.bytes().await?;
    let temporary = path.with_extension(format!("cloudreve-download-{}", Uuid::new_v4()));
    tokio::fs::write(&temporary, content).await?;
    tokio::fs::rename(&temporary, path).await?;
    Ok(())
}

async fn get_remote_file(state: &BridgeState, document: &OfficeManagedDocument) -> anyhow::Result<Option<FileResponse>> {
    let mount = state
        .drive_manager
        .get_drive(&document.drive_id)
        .await
        .ok_or_else(|| anyhow::anyhow!("Cloudreve account is no longer available"))?;
    Ok(mount
        .cr_client
        .get_file_info(&GetFileInfoService {
            uri: Some(document.remote_uri.clone()),
            id: None,
            extended: Some(true),
            folder_summary: None,
        })
        .await
        .ok())
}

async fn managed_document(state: &BridgeState, path: &Path) -> Option<OfficeManagedDocument> {
    state.documents.lock().await.documents.get(&path_key(path)).cloned()
}

fn build_file_response(path: PathBuf, remote_file: Option<FileResponse>, document: Option<OfficeManagedDocument>) -> OfficeFileResponse {
    let metadata = fs::metadata(&path).ok();
    let name = remote_file.as_ref().map(|file| file.name.clone()).or_else(|| path.file_name().and_then(|name| name.to_str()).map(ToOwned::to_owned)).unwrap_or_else(|| "未命名文档".to_string());
    let size = remote_file.as_ref().map(|file| file.size).or_else(|| metadata.as_ref().map(|metadata| metadata.len() as i64)).unwrap_or_default();
    let updated_at = remote_file.as_ref().map(|file| file.updated_at.clone()).or_else(|| metadata.and_then(|metadata| metadata.modified().ok()).map(|time| DateTime::<Utc>::from(time).to_rfc3339())).unwrap_or_default();
    let current = remote_file.as_ref().and_then(|file| file.primary_entity.clone());
    let versions = remote_file
        .and_then(|file| file.extended_info)
        .and_then(|info| info.entities)
        .unwrap_or_default()
        .into_iter()
        .filter(|entity| entity.entity_type == file_type::FILE)
        .map(|entity| OfficeVersion {
            current: current.as_deref() == Some(entity.id.as_str()),
            id: entity.id,
            size: entity.size,
            created_at: entity.created_at,
        })
        .collect();
    OfficeFileResponse {
        name,
        path: path.to_string_lossy().into_owned(),
        size,
        updated_at,
        current_version: current,
        versions,
        managed: document.is_some(),
        external: document.is_some(),
        remote_uri: document.as_ref().map(|document| document.remote_uri.clone()),
        last_synced_at: document.as_ref().and_then(|document| document.last_synced_at.clone()),
        last_error: document.and_then(|document| document.last_error),
        document_open: office_lock_path(&path).is_some_and(|lock| lock.exists()),
    }
}

async fn finish_document_session(state: &BridgeState, path: &Path, force: bool) -> anyhow::Result<()> {
    let mut registry = state.documents.lock().await;
    if let Some(document) = registry.documents.get_mut(&path_key(path)) {
        if force || document.active_session.is_some() {
            document.active_session = None;
            save_registry(&state.registry_path, &registry)?;
        }
    }
    Ok(())
}

async fn update_document_error(state: &BridgeState, path: &Path, error: String) -> anyhow::Result<()> {
    let mut registry = state.documents.lock().await;
    if let Some(document) = registry.documents.get_mut(&path_key(path)) {
        document.last_error = Some(error.clone());
        save_registry(&state.registry_path, &registry)?;
    }
    Err(anyhow::anyhow!(error))
}

fn office_registry_path() -> anyhow::Result<PathBuf> {
    let home = dirs::home_dir().ok_or_else(|| anyhow::anyhow!("Failed to locate the user profile"))?;
    Ok(home.join(".cloudreve").join(OFFICE_SYNC_FILE))
}

fn load_registry(path: &Path) -> OfficeSyncRegistry {
    fs::read_to_string(path)
        .ok()
        .and_then(|content| serde_json::from_str(&content).ok())
        .unwrap_or_default()
}

fn save_registry(path: &Path, registry: &OfficeSyncRegistry) -> anyhow::Result<()> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    fs::write(path, serde_json::to_string_pretty(registry)?)?;
    Ok(())
}

fn office_lock_path(path: &Path) -> Option<PathBuf> {
    let name = path.file_name()?.to_str()?;
    Some(path.with_file_name(format!("~${name}")))
}

fn path_key(path: &Path) -> String {
    path.canonicalize()
        .unwrap_or_else(|_| path.to_path_buf())
        .to_string_lossy()
        .to_ascii_lowercase()
}

fn system_time_millis(time: SystemTime) -> i64 {
    time.duration_since(UNIX_EPOCH).unwrap_or_default().as_millis() as i64
}

fn is_within_root(path: &str, root: &str) -> bool {
    let Ok(path) = CrUri::new(path) else { return false; };
    let Ok(root) = CrUri::new(root) else { return false; };
    path.base(true) == root.base(true)
}

struct BridgeError(anyhow::Error);

impl From<anyhow::Error> for BridgeError {
    fn from(error: anyhow::Error) -> Self { Self(error) }
}

impl IntoResponse for BridgeError {
    fn into_response(self) -> axum::response::Response {
        let status = if self.0.to_string().contains("still open") { StatusCode::CONFLICT } else { StatusCode::BAD_REQUEST };
        (status, Json(serde_json::json!({"error": self.0.to_string()}))).into_response()
    }
}
