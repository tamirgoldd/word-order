import type {
  CrossReferencePlan,
  NumberedTarget,
  ParagraphInventory,
  PlanWarning,
  RepairPlan,
  SchemeProfile
} from "./types.js";

interface PlanInputs {
  sourceFingerprint: string;
  inventory: ParagraphInventory[];
  profile: SchemeProfile;
  targets: NumberedTarget[];
  numbering: RepairPlan["numbering"];
  anomalies: RepairPlan["anomalies"];
  trackedChanges: boolean;
}

function canonicalSection(value: string): string {
  const match = value.match(/^(\d+(?:\.\d+)*)(.*)$/);
  if (!match) return value.toLowerCase();
  const numbers = (match[1] ?? "").split(".").map((part) => String(Number(part))).join(".");
  return `${numbers}${(match[2] ?? "").toLowerCase()}`;
}

function canonicalArticle(value: string): string {
  if (/^\d+$/.test(value)) return String(Number(value));
  const romanValues: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  const roman = value.toUpperCase();
  let total = 0;
  for (let index = 0; index < roman.length; index += 1) {
    const current = romanValues[roman[index] ?? ""] ?? 0;
    const next = romanValues[roman[index + 1] ?? ""] ?? 0;
    total += current < next ? -current : current;
  }
  return String(total || value.toLowerCase());
}

function targetIndex(inputs: PlanInputs): Map<string, NumberedTarget> {
  const result = new Map<string, NumberedTarget>();
  const changes = new Map(inputs.numbering.map((change) => [change.paragraphIndex, change]));
  for (const target of inputs.targets) {
    result.set(target.semanticKey, target);
    const change = changes.get(target.paragraphIndex);
    if (!change) continue;
    const old = change.oldNumber.replace(/^(?:ARTICLE|Article|Section|section)\s+/, "");
    if (target.level === 0) result.set(`article:${canonicalArticle(old)}`, target);
    if (target.level >= 1) result.set(`section:${canonicalSection(old)}`, target);
  }
  return result;
}

function paragraphSectionByIndex(inputs: PlanInputs): Map<number, string> {
  const sections = new Map<number, string>();
  let current = "";
  for (const target of inputs.targets) {
    if (target.level === 1) current = target.semanticKey;
    sections.set(target.paragraphIndex, current);
  }
  let latest = "";
  const result = new Map<number, string>();
  for (const paragraph of inputs.inventory) {
    if (sections.has(paragraph.index)) latest = sections.get(paragraph.index) ?? latest;
    result.set(paragraph.index, latest);
  }
  return result;
}

function addReference(
  references: CrossReferencePlan[],
  warnings: PlanWarning[],
  targetMap: Map<string, NumberedTarget>,
  paragraphIndex: number,
  display: string,
  start: number,
  key: string
): void {
  const target = targetMap.get(key);
  if (target) {
    references.push({
      paragraphIndex,
      display,
      start,
      end: start + display.length,
      status: "resolved",
      targetParagraphIndex: target.paragraphIndex,
      targetNumber: target.number,
      bookmarkName: target.bookmarkName,
      reason: null
    });
  } else {
    const reason = `No numbered paragraph resolves ${display}.`;
    references.push({
      paragraphIndex,
      display,
      start,
      end: start + display.length,
      status: "unresolved",
      targetParagraphIndex: null,
      targetNumber: null,
      bookmarkName: null,
      reason
    });
    warnings.push({ code: "broken-reference", paragraphIndex, message: reason });
  }
}

function detectReferences(inputs: PlanInputs): { references: CrossReferencePlan[]; warnings: PlanWarning[] } {
  const references: CrossReferencePlan[] = [];
  const warnings: PlanWarning[] = [];
  const targetMap = targetIndex(inputs);
  const currentSections = paragraphSectionByIndex(inputs);
  const sectionPhrase = /\bSections?\s+((?:\d+(?:\.\d+)*(?:\([A-Za-z0-9]+\))*)(?:\s*(?:,|and|through|to)\s*\d+(?:\.\d+)*(?:\([A-Za-z0-9]+\))*)*)/gi;
  const articlePhrase = /\bArticles?\s+((?:[IVXLCDM]+|\d+)(?:\s*(?:,|and|through|to)\s*(?:[IVXLCDM]+|\d+))*)/gi;
  const clausePhrase = /\bclauses?\s+(\([A-Za-z0-9]+\)(?:\([A-Za-z0-9]+\))?)/gi;

  for (const paragraph of inputs.inventory) {
    if (paragraph.hasRefField) continue;
    const baseOffset = paragraph.manual?.stripLength ?? 0;
    const searchableText = paragraph.text.slice(baseOffset);
    for (const phrase of searchableText.matchAll(sectionPhrase)) {
      const captured = phrase[1] ?? "";
      const captureStart = baseOffset + (phrase.index ?? 0) + phrase[0].indexOf(captured);
      for (const number of captured.matchAll(/\d+(?:\.\d+)*(?:\([A-Za-z0-9]+\))*/g)) {
        const display = number[0];
        addReference(references, warnings, targetMap, paragraph.index, display, captureStart + (number.index ?? 0), `section:${canonicalSection(display)}`);
      }
    }
    for (const phrase of searchableText.matchAll(articlePhrase)) {
      const captured = phrase[1] ?? "";
      const captureStart = baseOffset + (phrase.index ?? 0) + phrase[0].indexOf(captured);
      for (const number of captured.matchAll(/[IVXLCDM]+|\d+/gi)) {
        const display = number[0];
        addReference(references, warnings, targetMap, paragraph.index, display, captureStart + (number.index ?? 0), `article:${canonicalArticle(display)}`);
      }
    }
    for (const phrase of searchableText.matchAll(clausePhrase)) {
      const display = phrase[1] ?? "";
      const section = currentSections.get(paragraph.index) ?? "";
      const key = `${section}${display.toLowerCase()}`;
      const start = baseOffset + (phrase.index ?? 0) + phrase[0].indexOf(display);
      addReference(references, warnings, targetMap, paragraph.index, display, start, key);
    }
  }
  return { references, warnings };
}

export function buildRepairPlan(inputs: PlanInputs): RepairPlan {
  const { references, warnings } = detectReferences(inputs);
  const resolved = references.filter((reference) => reference.status === "resolved");
  const status = inputs.trackedChanges
    ? "blocked"
    : inputs.anomalies.length > 0
      ? "needs-confirmation"
      : inputs.numbering.length > 0 || resolved.length > 0
        ? "ready"
        : "clean";
  return {
    version: 1,
    sourceFingerprint: inputs.sourceFingerprint,
    status,
    blockedReason: inputs.trackedChanges
      ? "Tracked changes are present. Accept or reject all changes in Word before scanning again."
      : null,
    profile: inputs.profile,
    inventory: inputs.inventory,
    targets: inputs.targets,
    numbering: inputs.numbering,
    crossReferences: references,
    anomalies: inputs.anomalies,
    warnings,
    summary: {
      paragraphs: inputs.inventory.length,
      numberedParagraphs: inputs.targets.length,
      numberingConversions: inputs.numbering.length,
      crossReferencesConverted: resolved.length,
      brokenReferences: references.filter((reference) => reference.status === "unresolved").length,
      anomalies: inputs.anomalies.length
    }
  };
}
