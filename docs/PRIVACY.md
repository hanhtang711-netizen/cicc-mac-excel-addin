# 隐私说明

本加载项在 Excel 的 Office.js 运行环境内读取用户当前选区，并在当前工作簿中创建原生图表或应用单元格格式。工作簿数据留在 Excel 内部处理。

首版明确不包含：

- 用户账户、登录或身份系统；
- telemetry、埋点、分析 SDK 或使用情况上报；
- 工作簿、选区、公式、数字格式、图表数据或文件上传；
- 数据 API、云端模型调用、版本检查或远程数据服务；
- 工作簿内容的控制台或服务器日志记录。

生产 HTTPS 静态站点只提供加载项自身的 HTML、JavaScript、CSS 和图标。页面还会从 Microsoft 的 `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` 加载 Office.js。静态站点和 Office.js CDN 会像普通网页服务一样收到资源请求所需的网络元数据（例如 IP、User-Agent 和请求时间），但本加载项不会在这些请求中附带工作簿数据。

manifest 使用 `ReadWriteDocument` 权限，以读取当前选区，并把图表/格式写回当前文档。该权限不是上传授权。加载项不会读取用户账户，也不会主动遍历、复制或发送整个工作簿。

用户可随时按 [INSTALL_MAC.md](./INSTALL_MAC.md) 删除 manifest 完成卸载。若未来增加账户、telemetry、上传、远程 API 或其他数据流，必须先更新本说明、完成安全与隐私审查，并重新取得用户授权。
