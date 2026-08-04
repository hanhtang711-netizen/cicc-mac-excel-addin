import { CHART_CATALOG } from "../charts/chartCatalog";
import type { ChartKind, FeedbackPort } from "../core/types";
import type { ChartService, ServiceResult } from "./chartService";
import type { TableService } from "./tableService";

type CommandHandler = (event: Office.AddinCommands.Event) => Promise<void>;

export interface CommandDependencies {
  charts: Pick<ChartService, "create">;
  tables: Pick<TableService, "formatStandard" | "applyZebra">;
  feedback: FeedbackPort;
}

export function createCommandHandlers(
  dependencies: CommandDependencies,
): Record<string, CommandHandler> {
  const handlers: Record<string, CommandHandler> = {};
  const catalogEntries = Object.entries(CHART_CATALOG) as Array<
    [ChartKind, (typeof CHART_CATALOG)[ChartKind]]
  >;

  for (const [kind, { actionId }] of catalogEntries) {
    handlers[actionId] = (event) =>
      runCommand(event, () => dependencies.charts.create(kind), dependencies.feedback);
  }

  handlers.formatCiccTable = (event) =>
    runCommand(event, () => dependencies.tables.formatStandard(), dependencies.feedback);
  handlers.applyZebraStripe = (event) =>
    runCommand(event, () => dependencies.tables.applyZebra(), dependencies.feedback);

  return handlers;
}

export function registerCommandHandlers(dependencies: CommandDependencies): void {
  for (const [actionId, handler] of Object.entries(createCommandHandlers(dependencies))) {
    Office.actions.associate(actionId, handler);
  }
}

async function runCommand(
  event: Office.AddinCommands.Event,
  operation: () => Promise<ServiceResult>,
  feedback: FeedbackPort,
): Promise<void> {
  try {
    const result = await operation();
    if (result.warnings.length > 0) {
      await feedback.showWarnings(result.warnings);
    }
  } catch (error) {
    try {
      await feedback.showError(error);
    } catch {
      // Feedback is best-effort; a failed dialog must not leave the command running.
    }
  } finally {
    event.completed();
  }
}
