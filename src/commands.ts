import { ChartService } from "./app/chartService";
import {
  registerCommandHandlers,
  type CommandDependencies,
} from "./app/commandHandlers";
import { TableService } from "./app/tableService";
import { DialogFeedback } from "./app/userFeedback";
import { getCapabilities, type RequirementChecker } from "./office/capability";
import { ExcelGateway } from "./office/excelGateway";

export function createCommandDependencies(requirements: RequirementChecker): CommandDependencies {
  const gateway = new ExcelGateway();
  const capabilities = getCapabilities(requirements);
  return {
    charts: new ChartService(gateway, capabilities),
    tables: new TableService(gateway, capabilities),
    feedback: new DialogFeedback(),
  };
}

Office.onReady(() => {
  registerCommandHandlers(createCommandDependencies(Office.context.requirements));
});
