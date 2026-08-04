# 部署与发布

本加载项分为两种互不混用的运行形态：开发旁加载固定使用 `https://localhost:3000`；生产发布必须使用已获授权、长期固定的公网 HTTPS origin（例如 `https://excel-addon.example.com`）。生产 manifest 不得引用 localhost。

## 发布前提

- 已由部署负责人书面确认生产 origin；只接受无用户名密码、路径、query、fragment 的 HTTPS origin。
- 静态站点能在该 origin 的根路径提供文件，并使用有效、受信任的 TLS 证书。
- 站点允许加载 `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`。这是 Microsoft 的 Office.js CDN。
- 当前提交已经通过自动化 gate，并附有填写完成的 Mac Excel 验收记录（由 [TEST_CHECKLIST.md](./TEST_CHECKLIST.md) 复制后填写）；不要把未勾选的清单模板当作验收证据，也不要从未验收的工作区直接发布。

## 生成 release bundle

在仓库根目录执行：

```bash
npm ci
npm test
npm run typecheck
export ADDIN_BASE_URL="https://excel-addon.example.com"
export RELEASE_OUT="/absolute/authorized/output/cicc-addin-release"
npm run release -- --base-url "$ADDIN_BASE_URL" --out "$RELEASE_OUT"
./node_modules/.bin/office-addin-manifest validate "$RELEASE_OUT/manifest.production.xml"
```

`npm run release` 会先运行 Vite production build，再 fail-closed 地校验公网 URL、构建入口、四个 manifest 图标、manifest token，以及解析符号链接后的真实输出路径。导入打包模块不会自动执行 CLI。显式 `--out` 是唯一可能被替换的输出目录；不要把仓库根目录、`dist/`、`manifest/` 或它们的符号链接别名用作输出。替换已有输出时，打包器会先保留同目录备份；若新目录切换失败，会恢复上一版，再清理临时目录。

输出布局：

```text
cicc-addin-release/
├── site/                         # 将此目录的内容发布到 HTTPS origin 根路径
│   ├── taskpane.html
│   ├── commands.html
│   ├── feedback.html
│   └── assets/
└── manifest.production.xml       # 分发/旁加载此文件，不部署为网页入口
```

发布前检查 `manifest.production.xml` 不含 `{{BASE_URL}}`，且所有 `SourceLocation`、图标和命令 URL 均指向同一个授权 origin。

## 静态托管与缓存

1. 将 `site/` 的内容先上传到同一 origin 下的临时版本目录，完成文件校验后再原子切换流量或站点指针；不要逐文件覆盖正在服务的版本。
2. Vite 带 hash 的 `assets/*` 可设置 `Cache-Control: public, max-age=31536000, immutable`。
3. `taskpane.html`、`commands.html` 和 `feedback.html` 应设置 `Cache-Control: no-cache`（或很短的 TTL），以便它们及时引用新 hash 资源。
4. 确认三个 HTML、全部 assets 和图标返回 `200`，MIME type 正确，且没有 HTTP 跳转、混合内容或证书错误。
5. 保存本次完整 bundle、commit SHA、发布时间和 manifest 校验结果，以便回滚。

静态 host 与 Office.js CDN 只接收正常的静态资源请求；加载项没有把工作簿值、公式、选区或图表数据发往这些服务。详见 [PRIVACY.md](./PRIVACY.md)。

## manifest 更新

- origin 未变化且仅更新前端代码时，保持同一 manifest `Id` 和固定 URL；发布新的 `site/` bundle 即可，随后完全退出并重启 Excel 验证。
- origin 变化、资源 URL 变化或 manifest 功能变化时，使用新的 origin 重新运行 release packager，执行官方 validator，并重新分发生成的 production manifest。
- 若修改 manifest 的发布版本，递增 `<Version>`，不要更换稳定的 `<Id>`，除非明确要创建另一个加载项。
- 开发 `manifest/manifest.dev.xml` 只允许 `https://localhost:3000`，不得作为生产分发文件。

## 更新与回滚

更新时保留上一版完整 `site/` bundle。出现故障时：

1. 停止继续发布；记录当前 commit 和症状。
2. 将站点原子切回上一版 `site/`，保持生产 origin 不变。
3. 清理/失效三个 HTML 的缓存；hash assets 可继续共存。
4. 完全退出 Excel 后重启，复测 Ribbon、十种图表和两种表格命令。
5. 如果本次同时更改了 manifest，则恢复上一版已验证 manifest，并按 [INSTALL_MAC.md](./INSTALL_MAC.md) 的卸载/安装流程重新旁加载。

生产部署会改变外部状态。没有用户授权的具体 host 时，只能生成或测试本地临时 bundle，不能执行上传、DNS、TLS 或托管平台变更。
