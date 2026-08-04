export type AddinErrorCode =
  | "invalid_selection"
  | "selection_too_small"
  | "unsupported_layout"
  | "pie_requires_one_series"
  | "scatter_requires_numeric_x"
  | "unsupported_api"
  | "protected_sheet"
  | "excel_runtime_error";

export class AddinError extends Error {
  readonly code: AddinErrorCode;
  readonly details: unknown;

  constructor(code: AddinErrorCode, details: unknown = undefined) {
    super(code);
    this.name = "AddinError";
    this.code = code;
    this.details = details;
  }
}
