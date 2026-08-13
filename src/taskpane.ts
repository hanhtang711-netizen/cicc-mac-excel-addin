import { readAdvancedOptions } from "./app/advancedOptions";
import { ChartService, type ServiceResult } from "./app/chartService";
import { messageForError, messageForWarning } from "./app/userFeedback";
import { startAutoRunner } from "./bridge/autoRunner";
import { CHART_CATALOG } from "./charts/chartCatalog";
import { getCapabilities, type RequirementChecker } from "./office/capability";
import { ExcelGateway } from "./office/excelGateway";

type ChartCreator = Pick<ChartService, "create">;

export function initializeAdvancedChartPane(chartService: ChartCreator): void {
  const form = document.querySelector<HTMLFormElement>("#advanced-chart-form");
  const controls = document.querySelector<HTMLFieldSetElement>("#chart-form-controls");
  const status = document.querySelector<HTMLElement>("#status");
  if (form === null || controls === null || status === null) {
    return;
  }

  const kind = input(form, "kind", HTMLSelectElement);
  if (kind === null) {
    setStatus(status, "error", "高级图表窗格未能初始化，请关闭后重新打开。");
    return;
  }

  populateChartKinds(kind);
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
      showResult(status, result);
    } catch (error) {
      setStatus(status, "error", messageForError(error));
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

function showResult(status: HTMLElement, result: ServiceResult): void {
  if (result.warnings.length === 0) {
    setStatus(status, "success", "图表已生成。");
    return;
  }
  setStatus(status, "warning", messageForWarning(result.warnings[0] ?? ""));
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

export function createAdvancedChartService(
  requirements: RequirementChecker,
  gateway?: ExcelGateway,
): ChartService {
  const capabilities = getCapabilities(requirements);
  return new ChartService(gateway ?? new ExcelGateway(), capabilities);
}

/**
 * 自动嵌入文档：把当前工作簿标记为 AutoShowTaskpaneWithDocument。
 * 保存后，该工作簿每次打开都会自动弹出本任务窗格——这是 agent 自动
 * 唤醒通道（全自动桥）的文档侧地基。幂等，失败仅告警不影响使用。
 */
export function ensureAutoShowWithDocument(): void {
  try {
    const settings = Office.context.document.settings;
    if (settings.get("Office.AutoShowTaskpaneWithDocument") !== true) {
      settings.set("Office.AutoShowTaskpaneWithDocument", true);
      settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.warn("AutoShowTaskpaneWithDocument 保存失败", result.error);
        }
      });
    }
  } catch (error) {
    console.warn("ensureAutoShowWithDocument 失败", error);
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
    ensureAutoShowWithDocument();
    const gateway = new ExcelGateway();
    const service = createAdvancedChartService(Office.context.requirements, gateway);
    initializeAdvancedChartPane(service);
    startAutoRunner(gateway, service);
  });
}
