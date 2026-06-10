import { invoke } from "@tauri-apps/api/core";

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
