import { invoke } from "@tauri-apps/api/core";
import type { ProcessTemplate } from "./process-docs/types";

interface ImportedProcessTemplate {
  template: ProcessTemplate;
  directory: string;
}

export async function readBinaryFile(path: string): Promise<Uint8Array> {
  const bytes = await invoke<number[]>("read_binary_file", { path });
  return Uint8Array.from(bytes);
}

export async function writeBinaryFile(path: string, bytes: Uint8Array): Promise<void> {
  await invoke("write_binary_file", { path, bytes: Array.from(bytes) });
}

export async function openSystemPath(path: string): Promise<void> {
  await invoke("open_system_path", { path });
}

export async function processTemplateUserDir(): Promise<string> {
  return invoke<string>("process_template_user_dir");
}

export async function processBuiltinTemplateDir(): Promise<string> {
  return invoke<string>("process_builtin_template_dir");
}

export async function loadUserProcessTemplates(): Promise<ProcessTemplate[]> {
  return invoke<ProcessTemplate[]>("load_user_process_templates");
}

export async function importProcessTemplate(sourcePath: string): Promise<ImportedProcessTemplate> {
  return invoke<ImportedProcessTemplate>("import_process_template", { sourcePath });
}
