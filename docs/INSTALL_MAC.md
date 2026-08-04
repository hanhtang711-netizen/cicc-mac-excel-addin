# Mac Excel 安装、更新与卸载

目标验收版本为 Microsoft Excel for Mac 16.111.2。加载项是 Office.js manifest，不需要启用 VBA 宏。

## 开发环境：localhost 旁加载

1. 在仓库根目录运行：

   ```bash
   npm install
   npm run manifest:dev
   npm run dev
   ```

2. 在 Safari 打开 `https://localhost:3000/taskpane.html`。开发服务器使用本地自签名证书；仅在确认地址确实为 `localhost:3000` 后接受证书警告。必须先让 macOS/浏览器信任该本地证书，否则 Excel 的 WebView 可能显示空白窗格。不要对生产证书绕过警告。
3. 完全退出 Excel（菜单栏 **Excel → 退出 Excel**，或 `⌘Q`），而不只是关闭工作簿。
4. 在 Finder 按 `⌘⇧G`，前往：

   ```text
   ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
   ```

5. 如果 `wef` 不存在，就在 `Documents` 下新建它。把仓库中的 `manifest/manifest.dev.xml` 复制到 `wef`；不要复制源工作簿。
6. 重新打开 Excel 和一个测试工作簿。确认 Ribbon 出现 **中金工具** 页签，再打开“高级生成”确认窗格能加载。

若页签未出现，依次确认 Excel 已完全退出重启、文件扩展名仍为 `.xml`、manifest 校验通过、localhost 服务仍在运行，且 `https://localhost:3000/taskpane.html` 没有证书错误。

## 生产环境安装

生产安装不运行 localhost 服务。部署负责人应提供已通过校验的 `manifest.production.xml`：

1. 用文本方式检查 manifest 中全部 URL 都指向获授权的固定 HTTPS host，且不含 `localhost`、`{{BASE_URL}}` 或未知 host。
2. 完全退出 Excel。
3. 将该 manifest 复制到同一个 `wef` 目录；如开发 manifest 仍在，先移走开发 manifest，避免出现两个同名页签。
4. 重启 Excel，确认 **中金工具** 页签、十种图表命令、表格命令和高级窗格可用。

不要在无法确认来源、证书或权限时安装 manifest。该 manifest 请求 `ReadWriteDocument`，用于读取当前选区并在当前工作簿创建图表/应用格式。

## 更新与重启验证

- 仅静态代码更新且生产 host/manifest 不变：完全退出并重启 Excel，使 WebView 重新加载已部署的 HTML；无需重复复制 manifest。
- manifest 本身更新：退出 Excel，替换 `wef` 中的对应 manifest，再重启。
- 每次更新后都要确认页签仍显示，并按 [TEST_CHECKLIST.md](./TEST_CHECKLIST.md) 执行核心回归。

## 卸载

1. 完全退出 Excel。
2. 在 Finder 打开 `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`。
3. 只删除本加载项的 `manifest.dev.xml` 或收到的 `manifest.production.xml`；不要删除其他加载项 manifest，也不要删除整个 Excel 容器。
4. 重启 Excel，确认 **中金工具** 页签不再出现。

如果卸载后仍显示旧页签，再次确认所有 Excel 进程均已退出，并检查 `wef` 中是否还有同一加载项（manifest `Id` 为 `1da9e671-a567-4ca2-8a15-36d2bceacf44`）的另一份 XML。不要用清空整个 Office 缓存作为常规卸载步骤。
