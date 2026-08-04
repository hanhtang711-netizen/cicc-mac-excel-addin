import { AddinError } from "../core/errors";
import type { SelectionSnapshot, TableFormatPlan } from "../core/types";
import type { AddinCapabilities } from "../office/capability";
import { buildStandardTablePlan, buildZebraPlan } from "../tables/tablePlanner";
import type { ServiceResult } from "./chartService";

export interface TableGateway {
  readSelection(): Promise<SelectionSnapshot>;
  applyTablePlan(plan: TableFormatPlan): Promise<void>;
}

export class TableService {
  constructor(
    private readonly gateway: TableGateway,
    private readonly capabilities: AddinCapabilities,
  ) {}

  async formatStandard(): Promise<ServiceResult> {
    return this.apply(buildStandardTablePlan);
  }

  async applyZebra(): Promise<ServiceResult> {
    return this.apply(buildZebraPlan);
  }

  private async apply(buildPlan: (snapshot: SelectionSnapshot) => TableFormatPlan): Promise<ServiceResult> {
    if (!this.capabilities.base) {
      throw new AddinError("unsupported_api");
    }
    const snapshot = await this.gateway.readSelection();
    validateTableSelection(snapshot);
    await this.gateway.applyTablePlan(buildPlan(snapshot));
    return { ok: true, warnings: [] };
  }
}

function validateTableSelection(snapshot: SelectionSnapshot): void {
  const address = localAddress(snapshot.address);
  const isEmpty = snapshot.rowCount < 1 || snapshot.columnCount < 1;
  const isDiscontiguous = /[,;]/.test(address);
  const isWholeRow = /^\$?\d+:\$?\d+$/.test(address);
  const isWholeColumn = /^\$?[A-Z]+:\$?[A-Z]+$/i.test(address);
  const isWholeSheet = /^\$?A\$?1:\$?XFD\$?1048576$/i.test(address);

  if (isEmpty || isDiscontiguous || isWholeRow || isWholeColumn || isWholeSheet) {
    throw new AddinError("invalid_selection");
  }
}

function localAddress(address: string): string {
  const separator = address.lastIndexOf("!");
  return separator < 0 ? address : address.slice(separator + 1);
}
