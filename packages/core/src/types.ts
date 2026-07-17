export type PlanStatus = "blocked" | "clean" | "needs-confirmation" | "ready";

export type TokenKind =
  | "article"
  | "section"
  | "decimal"
  | "lower-letter"
  | "upper-letter"
  | "lower-roman"
  | "upper-roman"
  | "decimal-item"
  | "ambiguous-i";

export interface ManualToken {
  raw: string;
  stripLength: number;
  kind: TokenKind;
  ordinal: number;
  components: number[];
  label: string;
  suffix: string;
}

export interface NativeNumbering {
  numId: number;
  level: number;
  rendered: string;
  counters: number[];
  format: string;
}

export interface ParagraphInventory {
  index: number;
  text: string;
  styleId: string | null;
  indentLeft: number | null;
  indentHanging: number | null;
  inTable: boolean;
  hasTrackedChanges: boolean;
  hasFields: boolean;
  hasRefField: boolean;
  manual: ManualToken | null;
  native: NativeNumbering | null;
}

export interface SchemeProfile {
  articleLabel: "ARTICLE" | "Article";
  articleFormat: "upperRoman" | "decimal" | "upperLetter";
  sectionLabel: "Section" | "section" | "none";
  sectionWidth: number;
  sectionSeparator: string;
  levelIndents: number[];
}

export interface NumberingChange {
  paragraphIndex: number;
  level: number;
  action: "convert-manual" | "renumber-native";
  oldNumber: string;
  newNumber: string;
  textPreview: string;
  confidence: "high" | "review";
  bookmarkName: string;
  semanticKey: string;
}

export interface NumberedTarget {
  paragraphIndex: number;
  level: number;
  number: string;
  bookmarkName: string;
  semanticKey: string;
}

export interface CrossReferencePlan {
  paragraphIndex: number;
  display: string;
  start: number;
  end: number;
  status: "resolved" | "unresolved";
  targetParagraphIndex: number | null;
  targetNumber: string | null;
  bookmarkName: string | null;
  reason: string | null;
}

export interface PlanAnomaly {
  code: "ambiguous-token" | "sequence-gap" | "parent-mismatch";
  paragraphIndex: number;
  message: string;
  oldNumber: string;
  proposedNumber: string | null;
}

export interface PlanWarning {
  code: "broken-reference" | "unsupported-reference-format";
  paragraphIndex: number;
  message: string;
}

export interface RepairPlan {
  version: 1;
  sourceFingerprint: string;
  status: PlanStatus;
  blockedReason: string | null;
  profile: SchemeProfile;
  inventory: ParagraphInventory[];
  targets: NumberedTarget[];
  numbering: NumberingChange[];
  crossReferences: CrossReferencePlan[];
  anomalies: PlanAnomaly[];
  warnings: PlanWarning[];
  summary: {
    paragraphs: number;
    numberedParagraphs: number;
    numberingConversions: number;
    crossReferencesConverted: number;
    brokenReferences: number;
    anomalies: number;
  };
}

export interface RepairOptions {
  allowAnomalies?: boolean;
}

export interface RepairResult {
  bytes: Uint8Array;
  plan: RepairPlan;
  changed: boolean;
}

export interface TextInvariant {
  before: string[];
  after: string[];
  equal: boolean;
}
