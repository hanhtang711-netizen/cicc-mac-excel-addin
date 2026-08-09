import { CHART_CATALOG } from "../charts/chartCatalog";
import { AddinError } from "../core/errors";
import type { ChartKind, ChartOptions } from "../core/types";

const orientations = ["auto", "columns", "rows"] as const;

export function readAdvancedOptions(form: HTMLFormElement): { kind: ChartKind; options: ChartOptions } {
  const kind = approvedChartKind(valueOf(form, "kind"));
  const orientation = approvedValue(valueOf(form, "orientation"), orientations, "invalid_orientation");
  return { kind, options: { orientation } };
}

function valueOf(form: HTMLFormElement, name: string): string {
  const element = form.elements.namedItem(name);
  if (!(element instanceof HTMLInputElement || element instanceof HTMLSelectElement)) {
    throw unsupportedLayout(`missing_${name}`);
  }
  return element.value;
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
