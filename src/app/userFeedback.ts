import type { AddinErrorCode } from "../core/errors";
import type { FeedbackPort } from "../core/types";

const ERROR_MESSAGES: Readonly<Record<AddinErrorCode, string>> = {
  invalid_selection: "请选择一个连续的矩形区域后重试。",
  selection_too_small: "图表数据至少需要两行两列。",
  unsupported_layout: "当前数据排列无法识别，请在高级生成中指定系列方向。",
  pie_requires_one_series: "饼图只能使用一个数值系列。请缩小选区，或在高级生成中指定一个系列。",
  scatter_requires_numeric_x: "散点图第一列必须是数值或日期型 X 轴。",
  unsupported_api: "当前 Excel 版本不支持此图表类型，请升级 Excel。",
  protected_sheet: "当前工作表受保护，无法写入图表或格式。",
  excel_runtime_error: "Excel 未能完成操作，请检查选区后重试。",
};

const WARNING_MESSAGES = {
  series_palette_reused: "图表已生成。系列超过六个，后续系列将循环使用中金配色。",
} as const;

export type AddinWarningCode = keyof typeof WARNING_MESSAGES;
type FeedbackCode = AddinErrorCode | AddinWarningCode;

export function messageForError(error: unknown): string {
  if (hasInvalidCustomSize(error)) {
    return "请填写有效的自定义宽度和高度（大于 0 的厘米数）。";
  }
  return ERROR_MESSAGES[errorCodeFor(error)];
}

export function messageForWarning(code: string): string {
  return isWarningCode(code)
    ? WARNING_MESSAGES[code]
    : "图表已生成，但有非阻塞提示。";
}

export class DialogFeedback implements FeedbackPort {
  constructor(private readonly origin: string = window.location.origin) {}

  async showError(error: unknown): Promise<void> {
    this.open(errorCodeFor(error));
  }

  async showWarnings(codes: string[]): Promise<void> {
    for (const code of codes) {
      if (isWarningCode(code)) {
        this.open(code);
      }
    }
  }

  private open(messageCode: FeedbackCode): void {
    const dialogUrl = new URL("/feedback.html", this.origin);
    if (dialogUrl.protocol !== "https:") {
      throw new Error("feedback_requires_https");
    }
    dialogUrl.searchParams.set("code", messageCode);
    Office.context.ui.displayDialogAsync(dialogUrl.toString(), { height: 24, width: 32 });
  }
}

function errorCodeFor(error: unknown): AddinErrorCode {
  if (typeof error === "object" && error !== null && "code" in error) {
    const code = error.code;
    if (typeof code === "string" && Object.hasOwn(ERROR_MESSAGES, code)) {
      return code as AddinErrorCode;
    }
  }
  return "excel_runtime_error";
}

function isWarningCode(code: string): code is AddinWarningCode {
  return Object.hasOwn(WARNING_MESSAGES, code);
}

function hasInvalidCustomSize(error: unknown): boolean {
  if (typeof error !== "object" || error === null || !("code" in error) || !("details" in error)) {
    return false;
  }
  const details = error.details;
  return error.code === "unsupported_layout" &&
    typeof details === "object" &&
    details !== null &&
    "reason" in details &&
    details.reason === "invalid_custom_size";
}
