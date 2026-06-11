import { message as showDialogMessage } from "@tauri-apps/plugin-dialog";

export async function showOperationError(error: unknown) {
  await showDialogMessage(error instanceof Error ? error.message : String(error), {
    title: "操作失败",
    kind: "error",
  });
}
