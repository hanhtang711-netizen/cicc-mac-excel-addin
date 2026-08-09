# 中金 Excel 图表工具（Mac）

适用于 Microsoft Excel for Mac 的 Office.js 加载项。它只处理当前工作簿中的选区，不会上传工作簿、公式或图表数据。

## 提供的功能

- **生成图表**：选中含表头的数据区域，选择柱形图、折线图、饼图、散点图等类型；图表会按中金固定尺寸、字体、配色和坐标轴样式生成。
- **格式化表格**：将首行设为深红表头（`#8A2626`），统一数据区的字号、对齐、行高和边框。
- **斑马纹**：独立为数据区添加浅灰交替行，不改变已有文本、数字格式或对齐。

图表系列配色依次为：`#640000`、`#B9B8A6`、`#3D889A`、`#BE995D`、`#646C86`、`#DD965D`、`#44546A`。表头色 `#8A2626` 不参与图表系列配色。

## 日常启动与使用（开发旁加载）

```bash
cd "/Users/sky/Documents/CICC plugin/.worktrees/cicc-mac-excel-addin"
npm run dev
```

保持该终端窗口运行；若显示 `Port 3000 is already in use`，说明服务已经启动，无需重复运行。

随后打开 Excel：

1. 选中数据区域（第一行是系列名称，第一列是日期/类别）。
2. 在功能区点击 **中金工具**。
3. 点击 **生成图表** 并选图表类型，或使用 **格式化表格**、**斑马纹**。
4. 代码更新后，新生成的图表会使用新样式；已有图表不会被自动改写。

首次安装、页签缺失、更新 manifest 以及卸载的完整步骤见 [docs/INSTALL_MAC.md](docs/INSTALL_MAC.md)。

## 开发验证

```bash
npm test
npm run typecheck
npm run build
```

生产发布的 HTTPS 部署要求见 [docs/DEPLOY.md](docs/DEPLOY.md)。
