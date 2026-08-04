import { ChartService } from "./app/chartService";
import {
  registerCommandHandlers,
  type CommandDependencies,
} from "./app/commandHandlers";
import { TableService } from "./app/tableService";
import { DialogFeedback } from "./app/userFeedback";
import { ExcelGateway } from "./office/excelGateway";

export function createCommandDependencies(): CommandDependencies {
  const gateway = new ExcelGateway();
  return {
    charts: new ChartService(gateway),
    tables: new TableService(gateway),
    feedback: new DialogFeedback(),
  };
}

Office.onReady(() => {
  registerCommandHandlers(createCommandDependencies());
});
