import { readAdvancedOptions } from "./app/advancedOptions";
import { ChartService, type ServiceResult } from "./app/chartService";
import { DialogFeedback } from "./app/userFeedback";
import { CHART_CATALOG } from "./charts/chartCatalog";
import { AddinError } from "./core/errors";
import type { FeedbackPort } from "./core/types";
import { ExcelGateway } from "./office/excelGateway";

type ChartCreator = Pick<ChartService, "create">;

export function initializeAdvancedChartPane(chartService: ChartCreator, feedback: FeedbackPort): void {
  const form = document.querySelector<HTMLFormElement>("#advanced-chart-form");
  const controls = document.querySelector<HTMLFieldSetElement>("#chart-form-controls");
  const status = document.querySelector<HTMLElement>("#status");
  if (form === null || controls === null || status === null) {
    return;
  }

  const kind = input(form, "kind", HTMLSelectElement);
  const sizePreset = input(form, "sizePreset", HTMLSelectElement);
  const customSizeFields = document.querySelector<HTMLElement>("#custom-size-fields");
  const trendlineField = document.querySelector<HTMLElement>("#trendline-field");
  const trendline = input(form, "addTrendline", HTMLInputElement);
  if (kind === null || sizePreset === null || customSizeFields === null || trendlineField === null || trendline === null) {
    setStatus(status, "error", "高级图表窗格未能初始化，请关闭后重新打开。");
    return;
  }

  populateChartKinds(kind);
  const updateConditionalFields = (): void => {
    customSizeFields.hidden = sizePreset.value !== "custom";
    const isScatter = kind.value === "scatterTrend";
    trendlineField.hidden = !isScatter;
    if (!isScatter) {
      trendline.checked = false;
    }
  };
  sizePreset.addEventListener("change", updateConditionalFields);
  kind.addEventListener("change", updateConditionalFields);
  updateConditionalFields();
  controls.disabled = false;

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const submitButton = form.querySelector<HTMLButtonElement>('button[type="submit"]');
    if (submitButton === null) {
      return;
    }
    submitButton.disabled = true;
    setStatus(status, "loading", "正在生成图表…");
    try {
      const request = readAdvancedOptions(form);
      const result = await chartService.create(request.kind, request.options);
      await showResult(status, result, feedback);
    } catch (error) {
      setStatus(status, "error", errorMessage(error));
      await bestEffort(() => feedback.showError(error));
    } finally {
      submitButton.disabled = false;
    }
  });
}

function populateChartKinds(select: HTMLSelectElement): void {
  select.replaceChildren();
  for (const [kind, chart] of Object.entries(CHART_CATALOG)) {
    const option = document.createElement("option");
    option.value = kind;
    option.textContent = chart.label;
    select.append(option);
  }
}

async function showResult(status: HTMLElement, result: ServiceResult, feedback: FeedbackPort): Promise<void> {
  if (result.warnings.length === 0) {
    setStatus(status, "success", "图表已生成。");
    return;
  }
  setStatus(status, "warning", "图表已生成，但部分系列将重复使用配色。");
  await bestEffort(() => feedback.showWarnings(result.warnings));
}

function input<T extends HTMLInputElement | HTMLSelectElement>(
  form: HTMLFormElement,
  name: string,
  constructor: { new (): T },
): T | null {
  const element = form.elements.namedItem(name);
  return element instanceof constructor ? element : null;
}

function setStatus(status: HTMLElement, state: "loading" | "success" | "warning" | "error", message: string): void {
  status.dataset.state = state;
  status.textContent = message;
}

function errorMessage(error: unknown): string {
  if (error instanceof AddinError && error.code === "unsupported_layout" &&
    (error.details as { reason?: string } | undefined)?.reason === "invalid_custom_size") {
    return "请填写有效的自定义宽度和高度（大于 0 的厘米数）。";
  }
  return "生成图表失败。请检查当前选区和设置后重试。";
}

async function bestEffort(operation: () => Promise<void>): Promise<void> {
  try {
    await operation();
  } catch {
    // The visible status is the primary feedback channel; dialog feedback is supplementary.
  }
}

if (typeof Office !== "undefined") {
  Office.onReady((info) => {
    if (info.host !== Office.HostType.Excel) {
      const status = document.querySelector<HTMLElement>("#status");
      if (status !== null) {
        setStatus(status, "error", "请在 Excel 中打开此图表工具。");
      }
      return;
    }
    initializeAdvancedChartPane(new ChartService(new ExcelGateway()), new DialogFeedback());
  });
}
