use base64::{engine::general_purpose, Engine as _};
use std::io::Write;
use std::path::{Path, PathBuf};

#[tauri::command]
pub fn read_binary_file(path: String) -> Result<Vec<u8>, String> {
    std::fs::read(path).map_err(|error| error.to_string())
}

#[tauri::command]
pub fn read_binary_file_base64(path: String) -> Result<String, String> {
    let bytes = std::fs::read(path).map_err(|error| error.to_string())?;
    Ok(general_purpose::STANDARD.encode(bytes))
}

#[tauri::command]
pub fn write_binary_file(path: String, bytes: Vec<u8>) -> Result<(), String> {
    write_binary_file_verified(Path::new(&path), &bytes)
}

#[tauri::command]
pub fn write_binary_file_base64(
    path: String,
    content_base64: String,
    expected_len: usize,
) -> Result<(), String> {
    let bytes = general_purpose::STANDARD
        .decode(content_base64)
        .map_err(|error| format!("文件内容编码无效：{error}"))?;
    if bytes.len() != expected_len {
        return Err(format!(
            "文件写入前大小校验失败：期望 {} 字节，实际 {} 字节",
            expected_len,
            bytes.len()
        ));
    }

    write_binary_file_verified(Path::new(&path), &bytes)
}

pub fn write_binary_file_verified(path: &Path, bytes: &[u8]) -> Result<(), String> {
    let parent = path
        .parent()
        .ok_or_else(|| "输出文件路径无效，无法识别父目录".to_string())?;
    std::fs::create_dir_all(parent).map_err(|error| error.to_string())?;

    let temp_path = unique_temp_output_path(path);
    let write_result = (|| -> Result<(), String> {
        let mut file = std::fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&temp_path)
            .map_err(|error| format!("无法创建临时输出文件：{error}"))?;
        file.write_all(bytes)
            .map_err(|error| format!("写入临时输出文件失败：{error}"))?;
        file.sync_all()
            .map_err(|error| format!("同步临时输出文件失败：{error}"))?;
        drop(file);

        match std::fs::rename(&temp_path, path) {
            Ok(()) => {}
            Err(rename_error) if path.exists() => {
                std::fs::remove_file(path)
                    .map_err(|remove_error| format!("替换已有输出文件失败：{remove_error}"))?;
                std::fs::rename(&temp_path, path).map_err(|error| {
                    format!("替换输出文件失败：{error}；首次重命名错误：{rename_error}")
                })?;
            }
            Err(error) => return Err(format!("移动临时输出文件失败：{error}")),
        }
        if let Ok(parent_dir) = std::fs::File::open(parent) {
            let _ = parent_dir.sync_all();
        }

        let written = std::fs::read(path).map_err(|error| format!("回读输出文件失败：{error}"))?;
        if written.len() != bytes.len() {
            return Err(format!(
                "文件写入后大小校验失败：期望 {} 字节，实际 {} 字节",
                bytes.len(),
                written.len()
            ));
        }
        if written != bytes {
            return Err("文件写入后内容校验失败，输出文件可能未完整写入。请检查外置硬盘连接或改用本机磁盘后重试。".to_string());
        }

        Ok(())
    })();

    if write_result.is_err() && temp_path.exists() {
        let _ = std::fs::remove_file(&temp_path);
    }

    write_result
}

fn unique_temp_output_path(path: &Path) -> PathBuf {
    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    let file_name = path
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or("output");
    let pid = std::process::id();
    let nanos = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|value| value.as_nanos())
        .unwrap_or(0);
    parent.join(format!(".{file_name}.{pid}.{nanos}.tmp"))
}

#[tauri::command]
pub fn open_system_path(app: tauri::AppHandle, path: String) -> Result<(), String> {
    use tauri_plugin_opener::OpenerExt;

    app.opener()
        .open_path(path, None::<&str>)
        .map_err(|error| error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn verified_binary_write_replaces_existing_file_and_preserves_bytes() {
        let dir = std::env::temp_dir().join(format!(
            "archive-docx-client-write-test-{}",
            std::process::id()
        ));
        let path = dir.join("nested").join("output.xlsx");
        std::fs::create_dir_all(path.parent().unwrap()).unwrap();
        std::fs::write(&path, b"old bytes").unwrap();

        let bytes = b"PK\x03\x04test workbook bytes";
        write_binary_file_verified(&path, bytes).unwrap();

        assert_eq!(std::fs::read(&path).unwrap(), bytes);
        let _ = std::fs::remove_dir_all(dir);
    }
}
