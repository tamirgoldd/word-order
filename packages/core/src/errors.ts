export type WordOrderErrorCode =
  | "INVALID_DOCX"
  | "TRACKED_CHANGES"
  | "SOURCE_MISMATCH"
  | "CONFIRMATION_REQUIRED"
  | "UNSUPPORTED_DOCUMENT";

export class WordOrderError extends Error {
  readonly code: WordOrderErrorCode;

  constructor(code: WordOrderErrorCode, message: string) {
    super(message);
    this.name = "WordOrderError";
    this.code = code;
  }
}
