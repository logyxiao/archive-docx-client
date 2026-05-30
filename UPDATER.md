# 应用内更新配置

本项目已接入 Tauri v2 updater。用户需要先手动安装一次带 updater 的版本，之后就可以在应用内点击“检查更新”完成更新。

## GitHub Secrets

在 GitHub 仓库的 `Settings -> Secrets and variables -> Actions` 中添加：

- `TAURI_SIGNING_PRIVATE_KEY`：本地 `updater-private.key` 文件内容
- `TAURI_SIGNING_PRIVATE_KEY_PASSWORD`：本次生成密钥时密码为空，可先留空

`updater-private.key` 和 `updater-private.key.pub` 已加入 `.gitignore`，不要提交到仓库。

## 发布更新

推送 tag 会自动构建 Windows + macOS，并发布 GitHub Release：

```bash
git tag v0.2.0
git push origin v0.2.0
```

Release 中会包含安装包、签名文件和 `latest.json`。应用内“检查更新”会读取：

```text
https://github.com/logyxiao/archive-docx-client/releases/latest/download/latest.json
```
