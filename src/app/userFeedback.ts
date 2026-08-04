import { AddinError } from "../core/errors";
import type { FeedbackPort } from "../core/types";

export class DialogFeedback implements FeedbackPort {
  constructor(private readonly origin: string = window.location.origin) {}

  async showError(error: unknown): Promise<void> {
    this.open(error instanceof AddinError ? error.code : "excel_runtime_error");
  }

  async showWarnings(codes: string[]): Promise<void> {
    for (const code of codes) {
      this.open(code);
    }
  }

  private open(messageCode: string): void {
    const dialogUrl = new URL("/feedback.html", this.origin);
    dialogUrl.searchParams.set("code", messageCode);
    Office.context.ui.displayDialogAsync(dialogUrl.toString(), { height: 24, width: 32 });
  }
}
