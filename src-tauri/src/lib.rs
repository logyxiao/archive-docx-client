use serde::{Deserialize, Serialize};
use std::path::{Path, PathBuf};
use tauri::Manager;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct UserProcessTemplate {
    sequence: u32,
    kind: String,
    original_name: String,
    template_file: String,
    output_extension: String,
    user_template_path: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct ImportedProcessTemplate {
    template: UserProcessTemplate,
    directory: String,
}

#[tauri::command]
fn read_binary_file(path: String) -> Result<Vec<u8>, String> {
    std::fs::read(path).map_err(|error| error.to_string())
}

#[tauri::command]
fn write_binary_file(path: String, bytes: Vec<u8>) -> Result<(), String> {
    if let Some(parent) = std::path::Path::new(&path).parent() {
        std::fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }

    std::fs::write(path, bytes).map_err(|error| error.to_string())
}

#[tauri::command]
fn open_system_path(app: tauri::AppHandle, path: String) -> Result<(), String> {
    use tauri_plugin_opener::OpenerExt;

    app.opener()
        .open_path(path, None::<&str>)
        .map_err(|error| error.to_string())
}

#[tauri::command]
fn process_template_user_dir(app: tauri::AppHandle) -> Result<String, String> {
    let dir = process_template_dir(&app)?;
    std::fs::create_dir_all(&dir).map_err(|error| error.to_string())?;
    Ok(path_to_string(dir))
}

#[tauri::command]
fn process_builtin_template_dir() -> Result<String, String> {
    find_builtin_template_dir()
        .map(path_to_string)
        .ok_or_else(|| "未找到内置模板目录".to_string())
}

#[tauri::command]
fn load_user_process_templates(app: tauri::AppHandle) -> Result<Vec<UserProcessTemplate>, String> {
    let dir = process_template_dir(&app)?;
    sync_user_process_templates(&dir)
}

#[tauri::command]
fn import_process_template(
    app: tauri::AppHandle,
    source_path: String,
) -> Result<ImportedProcessTemplate, String> {
    let source = PathBuf::from(source_path);
    let source_file_name = source
        .file_name()
        .and_then(|value| value.to_str())
        .ok_or_else(|| "无法识别模板文件名".to_string())?
        .to_string();
    let extension = source
        .extension()
        .and_then(|value| value.to_str())
        .map(|value| value.to_ascii_lowercase())
        .ok_or_else(|| "仅支持导入 .docx 或 .xlsx 模板".to_string())?;

    if extension != "docx" && extension != "xlsx" {
        return Err("仅支持导入 .docx 或 .xlsx 模板".to_string());
    }

    let dir = process_template_dir(&app)?;
    std::fs::create_dir_all(&dir).map_err(|error| error.to_string())?;
    let dest = unique_template_path(&dir, &source_file_name);
    std::fs::copy(&source, &dest).map_err(|error| error.to_string())?;

    let mut templates = sync_user_process_templates(&dir)?;
    let dest_file_name = dest
        .file_name()
        .and_then(|value| value.to_str())
        .ok_or_else(|| "无法识别导入后的模板文件名".to_string())?
        .to_string();
    let template = UserProcessTemplate {
        sequence: 0,
        kind: extension.clone(),
        original_name: source_file_name,
        template_file: dest_file_name,
        output_extension: format!(".{}", extension),
        user_template_path: path_to_string(dest),
    };

    templates.retain(|item| item.user_template_path != template.user_template_path);
    templates.push(template.clone());
    templates.sort_by(|left, right| left.template_file.cmp(&right.template_file));
    write_user_process_templates(&dir, &templates)?;

    Ok(ImportedProcessTemplate {
        template,
        directory: path_to_string(dir),
    })
}

fn process_template_dir(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    app.path()
        .app_data_dir()
        .map(|dir| dir.join("process-templates"))
        .map_err(|error| error.to_string())
}

fn manifest_path(dir: &Path) -> PathBuf {
    dir.join("manifest.json")
}

fn read_user_process_templates(dir: &Path) -> Result<Vec<UserProcessTemplate>, String> {
    let path = manifest_path(dir);
    if !path.exists() {
        return Ok(Vec::new());
    }

    let content = std::fs::read_to_string(path).map_err(|error| error.to_string())?;
    serde_json::from_str(&content).map_err(|error| error.to_string())
}

fn sync_user_process_templates(dir: &Path) -> Result<Vec<UserProcessTemplate>, String> {
    std::fs::create_dir_all(dir).map_err(|error| error.to_string())?;
    let mut templates = read_user_process_templates(dir)?;
    let mut changed = false;
    let existing_paths = templates
        .iter()
        .map(|template| template.user_template_path.clone())
        .collect::<std::collections::HashSet<_>>();

    for entry in std::fs::read_dir(dir).map_err(|error| error.to_string())? {
        let entry = entry.map_err(|error| error.to_string())?;
        let path = entry.path();
        if !path.is_file() || existing_paths.contains(&path_to_string(path.clone())) {
            continue;
        }

        let Some(extension) = path
            .extension()
            .and_then(|value| value.to_str())
            .map(|value| value.to_ascii_lowercase())
        else {
            continue;
        };
        if extension != "docx" && extension != "xlsx" {
            continue;
        }

        let Some(file_name) = path.file_name().and_then(|value| value.to_str()) else {
            continue;
        };
        templates.push(UserProcessTemplate {
            sequence: 0,
            kind: extension.clone(),
            original_name: file_name.to_string(),
            template_file: file_name.to_string(),
            output_extension: format!(".{}", extension),
            user_template_path: path_to_string(path),
        });
        changed = true;
    }

    if changed {
        templates.sort_by(|left, right| left.template_file.cmp(&right.template_file));
        write_user_process_templates(dir, &templates)?;
    }

    Ok(templates)
}

fn write_user_process_templates(
    dir: &Path,
    templates: &[UserProcessTemplate],
) -> Result<(), String> {
    let content = serde_json::to_string_pretty(templates).map_err(|error| error.to_string())?;
    std::fs::write(manifest_path(dir), content).map_err(|error| error.to_string())
}

fn unique_template_path(dir: &Path, file_name: &str) -> PathBuf {
    let mut candidate = dir.join(file_name);
    if !candidate.exists() {
        return candidate;
    }

    let source = Path::new(file_name);
    let stem = source
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("template");
    let extension = source
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or("");

    for index in 1.. {
        let next_name = if extension.is_empty() {
            format!("{}-{}", stem, index)
        } else {
            format!("{}-{}.{}", stem, index, extension)
        };
        candidate = dir.join(next_name);
        if !candidate.exists() {
            return candidate;
        }
    }

    candidate
}

fn path_to_string(path: PathBuf) -> String {
    path.to_string_lossy().to_string()
}

fn find_builtin_template_dir() -> Option<PathBuf> {
    let candidates = [
        std::env::current_dir()
            .ok()?
            .join("public/templates/process-docs"),
        std::env::current_dir()
            .ok()?
            .parent()?
            .join("public/templates/process-docs"),
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .parent()?
            .join("public/templates/process-docs"),
    ];

    candidates.into_iter().find(|path| path.exists())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_process::init())
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .invoke_handler(tauri::generate_handler![
            read_binary_file,
            write_binary_file,
            open_system_path,
            process_template_user_dir,
            process_builtin_template_dir,
            load_user_process_templates,
            import_process_template
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
