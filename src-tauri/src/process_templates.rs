use serde::{Deserialize, Serialize};
use std::path::{Path, PathBuf};
use tauri::Manager;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct UserProcessTemplate {
    sequence: u32,
    kind: String,
    original_name: String,
    template_file: String,
    output_extension: String,
    user_template_path: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    display_name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    match_keywords: Option<Vec<String>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    match_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    template_module: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    category: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    enabled: Option<bool>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ImportedProcessTemplate {
    template: UserProcessTemplate,
    directory: String,
}

#[tauri::command]
pub fn process_template_user_dir(app: tauri::AppHandle) -> Result<String, String> {
    let dir = process_template_dir(&app)?;
    std::fs::create_dir_all(&dir).map_err(|error| error.to_string())?;
    Ok(path_to_string(dir))
}

#[tauri::command]
pub fn process_builtin_template_dir(app: tauri::AppHandle) -> Result<String, String> {
    let source = find_builtin_template_dir(&app).ok_or_else(|| "未找到内置模板目录".to_string())?;
    let mirror = app
        .path()
        .app_data_dir()
        .map(|dir| dir.join("builtin-process-templates"))
        .map_err(|error| error.to_string())?;

    sync_directory(&source, &mirror)?;
    Ok(path_to_string(mirror))
}

#[tauri::command]
pub fn load_user_process_templates(
    app: tauri::AppHandle,
) -> Result<Vec<UserProcessTemplate>, String> {
    let dir = process_template_dir(&app)?;
    sync_user_process_templates(&dir)
}

#[tauri::command]
pub fn import_process_template(
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
        display_name: Some(template_display_name(
            source
                .file_name()
                .and_then(|value| value.to_str())
                .unwrap_or(""),
        )),
        match_keywords: Some(infer_default_match_keywords(
            source
                .file_name()
                .and_then(|value| value.to_str())
                .unwrap_or(""),
        )),
        match_mode: Some("any".to_string()),
        template_module: Some("process".to_string()),
        category: None,
        enabled: Some(true),
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

#[tauri::command]
pub fn update_user_process_template(
    app: tauri::AppHandle,
    template: UserProcessTemplate,
) -> Result<UserProcessTemplate, String> {
    let dir = process_template_dir(&app)?;
    let requested_path = template.user_template_path.clone();
    let target = validate_user_template_path(&dir, &template.user_template_path)?;
    if !target.is_file() {
        return Err("未找到用户模板文件".to_string());
    }

    let mut templates = sync_user_process_templates(&dir)?;
    let target_path = path_to_string(target.clone());
    let mut normalized = normalize_user_template(template, &target)?;
    normalized.user_template_path = target_path.clone();

    let mut found = false;
    for item in &mut templates {
        if item.user_template_path == target_path || item.user_template_path == requested_path {
            *item = normalized.clone();
            found = true;
            break;
        }
    }

    if !found {
        templates.push(normalized.clone());
    }

    templates.sort_by(|left, right| left.template_file.cmp(&right.template_file));
    write_user_process_templates(&dir, &templates)?;
    Ok(normalized)
}

#[tauri::command]
pub fn delete_user_process_template(app: tauri::AppHandle, path: String) -> Result<(), String> {
    let dir = process_template_dir(&app)?;
    let target = validate_user_template_path(&dir, &path)?;
    let target_path = path_to_string(target.clone());

    let mut templates = sync_user_process_templates(&dir)?;
    templates
        .retain(|item| item.user_template_path != target_path && item.user_template_path != path);
    write_user_process_templates(&dir, &templates)?;

    if target.exists() {
        std::fs::remove_file(target).map_err(|error| error.to_string())?;
    }

    Ok(())
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
        let file_name = file_name.to_string();
        templates.push(UserProcessTemplate {
            sequence: 0,
            kind: extension.clone(),
            original_name: file_name.clone(),
            template_file: file_name.clone(),
            output_extension: format!(".{}", extension),
            user_template_path: path_to_string(path),
            display_name: Some(template_display_name(&file_name)),
            match_keywords: Some(infer_default_match_keywords(&file_name)),
            match_mode: Some("any".to_string()),
            template_module: Some("process".to_string()),
            category: None,
            enabled: Some(true),
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

fn validate_user_template_path(dir: &Path, path: &str) -> Result<PathBuf, String> {
    let canonical_dir = dir.canonicalize().map_err(|error| error.to_string())?;
    let candidate = PathBuf::from(path);
    let canonical_candidate = if candidate.exists() {
        candidate
            .canonicalize()
            .map_err(|error| error.to_string())?
    } else {
        let file_name = candidate
            .file_name()
            .ok_or_else(|| "模板路径无效".to_string())?;
        candidate
            .parent()
            .ok_or_else(|| "模板路径无效".to_string())?
            .canonicalize()
            .map(|parent| parent.join(file_name))
            .map_err(|error| error.to_string())?
    };

    if !canonical_candidate.starts_with(&canonical_dir) {
        return Err("只能管理用户模板目录内的模板".to_string());
    }

    Ok(canonical_candidate)
}

fn normalize_user_template(
    mut template: UserProcessTemplate,
    path: &Path,
) -> Result<UserProcessTemplate, String> {
    let file_name = path
        .file_name()
        .and_then(|value| value.to_str())
        .ok_or_else(|| "无法识别模板文件名".to_string())?
        .to_string();
    let extension = path
        .extension()
        .and_then(|value| value.to_str())
        .map(|value| value.to_ascii_lowercase())
        .ok_or_else(|| "仅支持 .docx 或 .xlsx 模板".to_string())?;

    if extension != "docx" && extension != "xlsx" {
        return Err("仅支持 .docx 或 .xlsx 模板".to_string());
    }

    template.kind = extension.clone();
    template.template_file = file_name.clone();
    template.output_extension = format!(".{}", extension);
    if template.original_name.trim().is_empty() {
        template.original_name = file_name.clone();
    }
    template.display_name = normalize_optional_string(template.display_name)
        .or_else(|| Some(template_display_name(&file_name)));
    template.match_keywords = template
        .match_keywords
        .map(normalize_keywords)
        .filter(|keywords| !keywords.is_empty());
    template.match_mode = normalize_match_mode(template.match_mode);
    template.template_module = normalize_template_module(template.template_module);
    template.category = normalize_category(template.category);
    template.enabled = Some(template.enabled.unwrap_or(true));

    Ok(template)
}

fn template_display_name(file_name: &str) -> String {
    Path::new(file_name)
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or(file_name)
        .trim()
        .to_string()
}

fn infer_default_match_keywords(file_name: &str) -> Vec<String> {
    let mut value = template_display_name(file_name);
    if let Some((left, right)) = value.split_once('、') {
        if !left.is_empty() && left.chars().all(|ch| ch.is_ascii_digit()) {
            value = right.trim().to_string();
        }
    }

    if let Some(index) = value
        .char_indices()
        .find(|(_, ch)| is_cjk(*ch))
        .map(|(index, _)| index)
    {
        value = value[index..].trim().to_string();
    }

    if let Some(index) = value.find("项目") {
        value = value[index + "项目".len()..].trim().to_string();
    }

    normalize_keywords(vec![value])
}

fn normalize_keywords(keywords: Vec<String>) -> Vec<String> {
    let mut result = Vec::new();
    for keyword in keywords {
        let trimmed = keyword.trim();
        if trimmed.is_empty() || result.iter().any(|item| item == trimmed) {
            continue;
        }
        result.push(trimmed.to_string());
    }
    result
}

fn normalize_optional_string(value: Option<String>) -> Option<String> {
    value.and_then(|item| {
        let trimmed = item.trim();
        if trimmed.is_empty() {
            None
        } else {
            Some(trimmed.to_string())
        }
    })
}

fn normalize_match_mode(value: Option<String>) -> Option<String> {
    match value.as_deref() {
        Some("all") => Some("all".to_string()),
        Some("any") | None => Some("any".to_string()),
        _ => Some("any".to_string()),
    }
}

fn normalize_template_module(value: Option<String>) -> Option<String> {
    match value.as_deref() {
        Some("switch-station") => Some("switch-station".to_string()),
        Some("collector-line") => Some("collector-line".to_string()),
        Some("process") | None => Some("process".to_string()),
        _ => Some("process".to_string()),
    }
}

fn normalize_category(value: Option<String>) -> Option<String> {
    const CATEGORIES: &[&str] = &[
        "start-report",
        "subunit-inspection-application",
        "summary-quality-acceptance",
        "division-inspection-application",
        "division-quality-acceptance",
        "subitem-inspection-application",
        "subitem-quality-acceptance",
        "inspection-lot-acceptance",
        "construction-record",
        "other",
    ];

    value.and_then(|item| {
        if CATEGORIES.contains(&item.as_str()) {
            Some(item)
        } else {
            None
        }
    })
}

fn is_cjk(ch: char) -> bool {
    ('\u{4e00}'..='\u{9fff}').contains(&ch)
}

fn path_to_string(path: PathBuf) -> String {
    path.to_string_lossy().to_string()
}

fn sync_directory(source: &Path, target: &Path) -> Result<(), String> {
    if target.exists() {
        std::fs::remove_dir_all(target).map_err(|error| error.to_string())?;
    }
    std::fs::create_dir_all(target).map_err(|error| error.to_string())?;
    copy_directory_contents(source, target)
}

fn copy_directory_contents(source: &Path, target: &Path) -> Result<(), String> {
    for entry in std::fs::read_dir(source).map_err(|error| error.to_string())? {
        let entry = entry.map_err(|error| error.to_string())?;
        let source_path = entry.path();
        let file_name = entry.file_name();
        if is_ignored_template_entry(&file_name) {
            continue;
        }

        let target_path = target.join(file_name);
        if source_path.is_dir() {
            std::fs::create_dir_all(&target_path).map_err(|error| error.to_string())?;
            copy_directory_contents(&source_path, &target_path)?;
        } else if source_path.is_file() {
            std::fs::copy(&source_path, &target_path).map_err(|error| error.to_string())?;
        }
    }

    Ok(())
}

fn is_ignored_template_entry(file_name: &std::ffi::OsStr) -> bool {
    let name = file_name.to_string_lossy();
    name.starts_with('.') || name.starts_with("~$")
}

fn find_builtin_template_dir(app: &tauri::AppHandle) -> Option<PathBuf> {
    let resource_dir = app.path().resource_dir().ok();
    let current_dir = std::env::current_dir().ok();
    let manifest_parent = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()?
        .to_path_buf();
    let mut candidates = Vec::new();

    if let Some(resource_dir) = resource_dir {
        candidates.push(resource_dir.join("builtin-process-templates"));
        candidates.push(resource_dir.join("public/templates/process-docs"));
        candidates.push(resource_dir.join("templates/process-docs"));
    }
    if let Some(current_dir) = current_dir {
        candidates.push(current_dir.join("public/templates/process-docs"));
        if let Some(parent) = current_dir.parent() {
            candidates.push(parent.join("public/templates/process-docs"));
        }
    }
    candidates.push(manifest_parent.join("public/templates/process-docs"));
    candidates.push(manifest_parent.join("dist/templates/process-docs"));

    candidates
        .into_iter()
        .find(|path| path.exists() && path.join("manifest.json").exists())
}
