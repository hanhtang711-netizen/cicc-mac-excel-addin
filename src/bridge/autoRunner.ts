import { ChartService } from "../app/chartService";
import { messageForError } from "../app/userFeedback";
import type { ChartKind, ChartOptions } from "../core/types";
import type { ExcelGateway } from "../office/excelGateway";

/**
 * 全自动桥（agent 驱动）：
 * 任务窗格启动后轮询本地桥服务，领取 agent 提交的图表指令并自动执行，
 * 完成后回报结果。指令结构见 schema/scripts/cicc_bridge/cicc_bridge.py。
 */

const BRIDGE_URL = "https://localhost:18768";
const POLL_MS = 2000;

interface BridgeJob {
  id: string;
  kind: ChartKind;
  sheet: string;
  range: string;
  options?: ChartOptions;
}

async function report(jobId: string, status: "done" | "error", message: string): Promise<void> {
  try {
    await fetch(`${BRIDGE_URL}/result`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id: jobId, status, message }),
    });
  } catch (error) {
    console.warn("桥回报失败", error);
  }
}

async function runJob(job: BridgeJob, gateway: ExcelGateway, chartService: ChartService): Promise<void> {
  try {
    await gateway.selectRange(job.sheet, job.range);
    const result = await chartService.create(job.kind, job.options ?? {});
    await report(job.id, "done", result.warnings.join("; "));
  } catch (error) {
    await report(job.id, "error", messageForError(error));
  }
}

export function startAutoRunner(gateway: ExcelGateway, chartService: ChartService): void {
  let busy = false;

  const tick = async (): Promise<void> => {
    if (busy) {
      return;
    }
    try {
      const response = await fetch(`${BRIDGE_URL}/next`);
      if (!response.ok) {
        return;
      }
      const job = (await response.json()) as BridgeJob | null;
      if (job === null) {
        return;
      }
      busy = true;
      try {
        await runJob(job, gateway, chartService);
      } finally {
        busy = false;
      }
    } catch {
      // 桥服务未启动或网络受限：静默，等待下一轮
    }
  };

  void tick();
  setInterval(() => void tick(), POLL_MS);
}
