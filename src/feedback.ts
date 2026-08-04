const messages: Readonly<Record<string, string>> = {
  invalid_selection: "请选择一个连续的矩形区域后重试。",
  selection_too_small: "图表数据至少需要两行两列。",
  unsupported_layout: "当前数据排列无法识别，请在高级生成中指定系列方向。",
  pie_requires_one_series: "饼图只能使用一个数值系列。请缩小选区，或在高级生成中指定一个系列。",
  scatter_requires_numeric_x: "散点图第一列必须是数值或日期型 X 轴。",
  unsupported_api: "当前 Excel 版本不支持此图表类型，请升级 Excel。",
  protected_sheet: "当前工作表受保护，无法写入图表或格式。",
  excel_runtime_error: "Excel 未能完成操作，请检查选区后重试。",
  series_palette_reused: "系列超过六个时将循环使用中金配色。",
};

const messageCode = new URLSearchParams(window.location.search).get("code") ?? "excel_runtime_error";
const messageElement = document.querySelector<HTMLElement>("#feedback-message");

if (messageElement !== null) {
  messageElement.textContent = messages[messageCode] ?? messages.excel_runtime_error ?? "操作未完成。";
}
