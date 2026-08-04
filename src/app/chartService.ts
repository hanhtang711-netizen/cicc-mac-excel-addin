import { buildChartPlan } from "../charts/chartPlanner";
import { buildChartStylePlan } from "../charts/chartStyle";
import { parseSelection } from "../core/selectionParser";
import type {
  ChartKind,
  ChartOptions,
  ChartPlan,
  ChartStylePlan,
  SelectionSnapshot,
} from "../core/types";

export interface ChartGateway {
  readSelection(): Promise<SelectionSnapshot>;
  createChart(plan: ChartPlan, style: ChartStylePlan): Promise<void>;
}

export interface ServiceResult {
  ok: true;
  warnings: string[];
}

export class ChartService {
  constructor(private readonly gateway: ChartGateway) {}

  async create(kind: ChartKind, options: ChartOptions = {}): Promise<ServiceResult> {
    const snapshot = await this.gateway.readSelection();
    const parsed = parseSelection(snapshot, options.orientation);
    const plan = buildChartPlan(parsed, kind, options);
    const style = buildChartStylePlan({
      seriesCount: plan.series.length,
      options,
      sourceFormat: parsed.numberFormats[parsed.headerRows]?.[1] ?? "General",
      categoryFormat: parsed.numberFormats[parsed.headerRows]?.[0] ?? "General",
      selectionColumn: snapshot.columnIndex,
      selectionColumnCount: snapshot.columnCount,
    });

    await this.gateway.createChart(plan, style);
    return { ok: true, warnings: style.warnings };
  }
}
