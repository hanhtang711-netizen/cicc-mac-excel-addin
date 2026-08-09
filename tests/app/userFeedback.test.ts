// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest";
import { DialogFeedback, messageForError, messageForWarning } from "../../src/app/userFeedback";
import { AddinError, type AddinErrorCode } from "../../src/core/errors";

const expectedErrors: ReadonlyArray<readonly [AddinErrorCode, string]> = [
  ["invalid_selection", "请选择一个连续的矩形区域后重试。"],
  ["selection_too_small", "图表数据至少需要两行两列。"],
  ["unsupported_layout", "当前数据排列无法识别，请在高级生成中指定系列方向。"],
  ["pie_requires_one_series", "饼图只能使用一个数值系列。请缩小选区，或在高级生成中指定一个系列。"],
  ["scatter_requires_numeric_x", "散点图第一列必须是数值或日期型 X 轴。"],
  ["unsupported_api", "当前 Excel 版本不支持此图表类型，请升级 Excel。"],
  ["protected_sheet", "当前工作表受保护，无法写入图表或格式。"],
  ["excel_runtime_error", "Excel 未能完成操作，请检查选区后重试。"],
];

describe("user feedback", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it.each(expectedErrors)("maps %s to actionable Chinese copy", (code, message) => {
    expect(messageForError({ code })).toBe(message);
  });

  it("maps unknown failures to the Excel runtime message", () => {
    expect(messageForError(new Error("internal details"))).toBe(
      "Excel 未能完成操作，请检查选区后重试。",
    );
  });

  it("explains palette reuse as a non-blocking warning", () => {
    expect(messageForWarning("series_palette_reused")).toBe(
      "图表已生成。系列超过六个，后续系列将循环使用中金配色。",
    );
  });

  it("opens a same-origin HTTPS page containing only the stable error code", async () => {
    const displayDialogAsync = vi.fn();
    vi.stubGlobal("Office", { context: { ui: { displayDialogAsync } } });
    const feedback = new DialogFeedback("https://addin.example.com/taskpane.html");

    await feedback.showError(new AddinError("invalid_selection", {
      values: [["confidential"]],
      worksheetName: "Secret Plan",
    }));

    expect(displayDialogAsync).toHaveBeenCalledWith(
      "https://addin.example.com/feedback.html?code=invalid_selection",
      { height: 24, width: 32 },
    );
  });
});
