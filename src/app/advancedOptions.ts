import { CHART_CATALOG } from "../charts/chartCatalog";
import { AddinError } from "../core/errors";
import type { ChartKind, ChartOptions } from "../core/types";

const sizePresets = ["small", "medium", "large", "custom"] as const;
const legendPositions = ["bottom", "top", "left", "right", "none"] as const;
const orientations = ["columns", "rows"] as const;

export function readAdvancedOptions(form: HTMLFormElement): { kind: ChartKind; options: ChartOptions } {
  const kind = approvedChartKind(valueOf(form, "kind"));
  const sizePreset = approvedValue(valueOf(form, "sizePreset"), sizePresets, "invalid_size_preset");
  const legendPosition = approvedValue(valueOf(form, "legendPosition"), legendPositions, "invalid_legend_position");
  const orientation = approvedValue(valueOf(form, "orientation"), orientations, "invalid_orientation");
  const title = valueOf(form, "title").trim();
  const options: ChartOptions = {
    title: title || undefined,
    sizePreset,
    legendPosition,
    orientation,
    showDataLabels: checked(form, "showDataLabels"),
    ...(kind === "scatterTrend" ? { addTrendline: checked(form, "addTrendline") } : {}),
  };

  if (sizePreset === "custom") {
    const widthCm = Number(valueOf(form, "widthCm"));
    const heightCm = Number(valueOf(form, "heightCm"));
    if (!Number.isFinite(widthCm) || widthCm <= 0 || !Number.isFinite(heightCm) || heightCm <= 0) {
      throw unsupportedLayout("invalid_custom_size");
    }
    options.widthCm = widthCm;
    options.heightCm = heightCm;
  }

  return { kind, options };
}

function valueOf(form: HTMLFormElement, name: string): string {
  const element = form.elements.namedItem(name);
  if (!(element instanceof HTMLInputElement || element instanceof HTMLSelectElement)) {
    throw unsupportedLayout(`missing_${name}`);
  }
  return element.value;
}

function checked(form: HTMLFormElement, name: string): boolean {
  const element = form.elements.namedItem(name);
  if (!(element instanceof HTMLInputElement) || element.type !== "checkbox") {
    throw unsupportedLayout(`missing_${name}`);
  }
  return element.checked;
}

function approvedChartKind(value: string): ChartKind {
  if (Object.hasOwn(CHART_CATALOG, value)) {
    return value as ChartKind;
  }
  throw unsupportedLayout("invalid_chart_kind");
}

function approvedValue<T extends string>(value: string, approved: readonly T[], reason: string): T {
  const result = approved.find((item) => item === value);
  if (result === undefined) {
    throw unsupportedLayout(reason);
  }
  return result;
}

function unsupportedLayout(reason: string): AddinError {
  return new AddinError("unsupported_layout", { reason });
}
