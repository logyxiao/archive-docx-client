mod file_io;
mod process_templates;

use file_io::{
    open_system_path, read_binary_file, read_binary_file_base64, write_binary_file,
    write_binary_file_base64,
};
use process_templates::{
    delete_user_process_template, import_process_template, load_user_process_templates,
    process_builtin_template_dir, process_template_user_dir, update_user_process_template,
};

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
            read_binary_file_base64,
            write_binary_file,
            write_binary_file_base64,
            open_system_path,
            process_template_user_dir,
            process_builtin_template_dir,
            load_user_process_templates,
            import_process_template,
            update_user_process_template,
            delete_user_process_template
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
