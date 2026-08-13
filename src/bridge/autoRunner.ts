import { ChartService } from "../app/chartService";
import { messageForError } from "../app/userFeedback";
import type { ChartKind, ChartOptions } from "../core/types";

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

/** 诊断版错误描述：携带 name/code/message/details（Office.js 原始错误），
 * 兜底 messageForError 的中文提示。details 序列化失败则降级。 */
function describeError(error: unknown): string {
  const friendly = messageForError(error);
  if (typeof error !== "object" || error === null) {
    return friendly;
  }
  const e = error as { name?: string; code?: string; message?: string; details?: unknown };
  let detail = "";
  if (e.details !== undefined) {
    try {
      detail = JSON.stringify(e.details);
    } catch {
      detail = "<unserializable>";
    }
  }
  return `${e.name ?? "Error"}: ${e.message ?? ""} [code=${e.code ?? ""}] [details=${detail}]（${friendly}）`;
}

async function runJob(job: BridgeJob, chartService: ChartService): Promise<void> {
  try {
    // 直读地址出图，不经过"当前选区"（select+读选区有 UI 竞态，260814）
    const result = await chartService.create(job.kind, job.options ?? {}, {
      sheet: job.sheet,
      address: job.range,
    });
    await report(job.id, "done", result.warnings.join("; "));
  } catch (error) {
    await report(job.id, "error", describeError(error));
  }
}

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

export function startAutoRunner(chartService: ChartService): void {
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
        await runJob(job, chartService);
        // 坑位（260814）：Excel Rich API 快速连续批次时，第二个批次的
        // worksheets.getItem 命中未刷新缓存抛 ItemNotFound（奇偶交替复现）。
        // 每条指令完成后冷却 1.5s 再取下一条。
        await sleep(1500);
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
