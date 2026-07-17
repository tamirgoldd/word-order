export type LegalDownErrorCode =
  | "INVALID_DOCX"
  | "TRACKED_CHANGES"
  | "SOURCE_MISMATCH"
  | "CONFIRMATION_REQUIRED"
  | "UNSUPPORTED_DOCUMENT";

export class LegalDownError extends Error {
  readonly code: LegalDownErrorCode;

  constructor(code: LegalDownErrorCode, message: string) {
    super(message);
    this.name = "LegalDownError";
    this.code = code;
  }
}
