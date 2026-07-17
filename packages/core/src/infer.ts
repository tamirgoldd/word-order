import type {
  ManualToken,
  NumberedTarget,
  NumberingChange,
  ParagraphInventory,
  PlanAnomaly,
  SchemeProfile,
  TokenKind
} from "./types.js";
import { isFlatSectionHeading } from "./formatting.js";
import { numberToLetters, numberToRoman, romanToNumber } from "./tokens.js";

interface StructuralState {
  article: number | null;
  sectionComponents: number[] | null;
  sectionParentArticle: number | null;
  letter: number | null;
  roman: number | null;
}

function mostFrequent<T>(values: T[], fallback: T): T {
  const counts = new Map<T, number>();
  for (const value of values) counts.set(value, (counts.get(value) ?? 0) + 1);
  return [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? fallback;
}

export function inferProfile(inventory: ParagraphInventory[]): SchemeProfile {
  const manual = inventory.flatMap((paragraph) => paragraph.manual ? [paragraph.manual] : []);
  const articles = manual.filter((token) => token.kind === "article");
  const sections = manual.filter((token) => token.kind === "section" || token.kind === "decimal");
  const flatSectionHeadings = inventory.filter(isFlatSectionHeading);
  const nativeFlatSections = inventory.filter((paragraph) => paragraph.native?.level === 0 && paragraph.styleId === "LDSection");
  const topLevelKind = articles.length === 0 && sections.length === 0 && (flatSectionHeadings.length >= 2 || nativeFlatSections.length > 0)
    ? "flat-section" as const
    : "article" as const;
  const indents = [0, 720, 1440, 2160].map((fallback, level) => {
    const values = inventory
      .filter((paragraph) => effectiveLevel(paragraph, topLevelKind) === level || paragraph.native?.level === level)
      .flatMap((paragraph) => paragraph.indentLeft === null ? [] : [paragraph.indentLeft]);
    return mostFrequent(values, fallback);
  });
  return {
    topLevelKind,
    articleLabel: mostFrequent(articles.map((token) => token.label === "Article" ? "Article" as const : "ARTICLE" as const), "ARTICLE"),
    articleFormat: mostFrequent(articles.map((token) => /^\d+$/.test(token.raw.replace(/\D/g, "")) ? "decimal" as const : "upperRoman" as const), "upperRoman"),
    sectionLabel: mostFrequent(sections.map((token) => token.kind === "decimal" ? "none" as const : token.label === "section" ? "section" as const : "Section" as const), "Section"),
    sectionWidth: mostFrequent(sections.map((token) => token.raw.match(/\.(\d+)/)?.[1]?.length ?? 2), 2),
    sectionSeparator: ".",
    levelIndents: indents
  };
}

function effectiveLevel(paragraph: ParagraphInventory, topLevelKind: SchemeProfile["topLevelKind"]): number | null {
  if (topLevelKind === "flat-section" && isFlatSectionHeading(paragraph)) return 0;
  return tokenLevel(paragraph.manual) ?? nativeLevel(paragraph);
}

export function tokenLevel(token: ManualToken | null): number | null {
  if (!token) return null;
  switch (token.kind) {
    case "article": return 0;
    case "section":
    case "decimal": return 1;
    case "lower-letter":
    case "upper-letter":
    case "decimal-item": return 2;
    case "lower-roman":
    case "upper-roman": return 3;
    case "ambiguous-i": return null;
  }
}

function nativeLevel(paragraph: ParagraphInventory): number | null {
  return paragraph.native ? Math.min(paragraph.native.level, 3) : null;
}

function actualOrdinal(paragraph: ParagraphInventory, level: number): number {
  if (paragraph.manual) return paragraph.manual.ordinal;
  return paragraph.native?.counters[level] ?? 1;
}

function actualComponents(paragraph: ParagraphInventory): number[] {
  if (paragraph.manual && (paragraph.manual.kind === "section" || paragraph.manual.kind === "decimal")) {
    return paragraph.manual.components;
  }
  if (paragraph.native) return paragraph.native.counters.slice(0, 2);
  return [1, 1];
}

function formatArticle(ordinal: number, profile: SchemeProfile): string {
  const value = profile.articleFormat === "decimal"
    ? String(ordinal)
    : profile.articleFormat === "upperLetter"
      ? numberToLetters(ordinal, true)
      : numberToRoman(ordinal);
  return `${profile.articleLabel} ${value}`;
}

function formatSection(components: number[], profile: SchemeProfile): string {
  const values = components.map((value, index) => index === components.length - 1 && components.length > 1
    ? String(value).padStart(profile.sectionWidth, "0")
    : String(value));
  const number = values.join(profile.sectionSeparator);
  return profile.sectionLabel === "none" ? number : `${profile.sectionLabel} ${number}`;
}

function formatLevel(kind: TokenKind | null, level: number, ordinal: number): string {
  if (level === 2) {
    if (kind === "upper-letter") return `(${numberToLetters(ordinal, true)})`;
    if (kind === "decimal-item") return `(${ordinal})`;
    return `(${numberToLetters(ordinal)})`;
  }
  if (kind === "upper-roman") return `(${numberToRoman(ordinal)})`;
  return `(${numberToRoman(ordinal).toLowerCase()})`;
}

function stripLabel(number: string): string {
  return number.replace(/^(?:ARTICLE|Article|Section|section)\s+/, "");
}

function canonicalSection(number: string): string {
  const stripped = stripLabel(number).replace(/\.$/, "");
  const match = stripped.match(/^(\d+(?:\.\d+)*)(.*)$/);
  if (!match) return stripped.toLowerCase();
  const components = (match[1] ?? "").split(".").map((part) => String(Number(part))).join(".");
  return `${components}${(match[2] ?? "").toLowerCase()}`;
}

function semanticKey(level: number, number: string, state: StructuralState, profile: SchemeProfile): string {
  if (level === 0 && profile.topLevelKind === "flat-section") return `section:${canonicalSection(number)}`;
  if (level === 0) return `article:${romanToNumber(stripLabel(number)) || stripLabel(number)}`;
  if (level === 1) return `section:${canonicalSection(number)}`;
  const base = state.sectionComponents ? `section:${state.sectionComponents.join(".")}` : "clause";
  if (level === 2) return `${base}${number}`;
  const letter = state.letter === null ? "" : `(${numberToLetters(state.letter)})`;
  return `${base}${letter}${number}`;
}

function preview(paragraph: ParagraphInventory): string {
  const withoutToken = paragraph.manual ? paragraph.text.slice(paragraph.manual.stripLength) : paragraph.text;
  return withoutToken.trim().slice(0, 120);
}

export function inferNumbering(
  inventory: ParagraphInventory[],
  profile: SchemeProfile
): { changes: NumberingChange[]; targets: NumberedTarget[]; anomalies: PlanAnomaly[] } {
  const changes: NumberingChange[] = [];
  const targets: NumberedTarget[] = [];
  const anomalies: PlanAnomaly[] = [];
  const state: StructuralState = { article: null, sectionComponents: null, sectionParentArticle: null, letter: null, roman: null };

  for (const paragraph of inventory) {
    const level = effectiveLevel(paragraph, profile.topLevelKind);
    if (paragraph.manual?.kind === "ambiguous-i") {
      anomalies.push({
        code: "ambiguous-token",
        paragraphIndex: paragraph.index,
        message: "The token (i) has no sequence context, so it could be a letter or a Roman numeral.",
        oldNumber: paragraph.manual.raw,
        proposedNumber: null
      });
      continue;
    }
    if (level === null) continue;

    const oldNumber = paragraph.manual?.raw ?? paragraph.native?.rendered ?? "";
    let newNumber = oldNumber;
    let proposedOrdinal = actualOrdinal(paragraph, level);
    let anomalyCode: PlanAnomaly["code"] | null = null;
    let anomalyMessage = "";

    if (level === 0) {
      const actual = actualOrdinal(paragraph, level);
      const expected = state.article === null ? actual : state.article + 1;
      proposedOrdinal = expected;
      if (actual !== expected) {
        anomalyCode = "sequence-gap";
        const label = profile.topLevelKind === "flat-section" ? "Section" : "Article";
        anomalyMessage = `${label} sequence jumps from ${state.article ?? "the start"} to ${actual}.`;
      }
      state.article = expected;
      state.sectionComponents = null;
      state.sectionParentArticle = expected;
      state.letter = null;
      state.roman = null;
      newNumber = profile.topLevelKind === "flat-section" ? `${expected}.` : formatArticle(expected, profile);
    } else if (level === 1) {
      const actual = actualComponents(paragraph);
      const components = [...actual];
      if (components.length === 1) components.unshift(state.article ?? 1);
      const newParent = state.sectionComponents === null || state.sectionParentArticle !== state.article;
      const expectedMajor = state.article ?? components[0] ?? 1;
      const expectedMinor = newParent
        ? components.at(-1) ?? 1
        : (state.sectionComponents?.at(-1) ?? 0) + 1;
      if ((components[0] ?? expectedMajor) !== expectedMajor) {
        anomalyCode = "parent-mismatch";
        anomalyMessage = `Section ${components.join(".")} does not match its Article ${expectedMajor} parent.`;
      } else if (!newParent && (components.at(-1) ?? expectedMinor) !== expectedMinor) {
        anomalyCode = "sequence-gap";
        anomalyMessage = `Section sequence expected ${expectedMinor} but found ${components.at(-1)}.`;
      }
      components[0] = expectedMajor;
      components[components.length - 1] = expectedMinor;
      state.sectionComponents = components;
      state.sectionParentArticle = state.article;
      state.letter = null;
      state.roman = null;
      newNumber = formatSection(components, profile);
    } else if (level === 2) {
      const actual = actualOrdinal(paragraph, level);
      const expected = state.letter === null ? actual : state.letter + 1;
      proposedOrdinal = expected;
      if (actual !== expected) {
        anomalyCode = "sequence-gap";
        anomalyMessage = `Clause sequence expected ${formatLevel(paragraph.manual?.kind ?? null, level, expected)} but found ${oldNumber}.`;
      }
      state.letter = expected;
      state.roman = null;
      newNumber = formatLevel(paragraph.manual?.kind ?? null, level, expected);
    } else {
      const actual = actualOrdinal(paragraph, level);
      const expected = state.roman === null ? actual : state.roman + 1;
      proposedOrdinal = expected;
      if (actual !== expected) {
        anomalyCode = "sequence-gap";
        anomalyMessage = `Subclause sequence expected ${formatLevel(paragraph.manual?.kind ?? null, level, expected)} but found ${oldNumber}.`;
      }
      state.roman = expected;
      newNumber = formatLevel(paragraph.manual?.kind ?? null, level, expected);
    }

    const bookmarkName = `_LDRef_${paragraph.index + 1}`;
    const key = semanticKey(level, newNumber, state, profile);
    targets.push({ paragraphIndex: paragraph.index, level, number: newNumber, bookmarkName, semanticKey: key });

    if (anomalyCode) {
      anomalies.push({
        code: anomalyCode,
        paragraphIndex: paragraph.index,
        message: anomalyMessage,
        oldNumber,
        proposedNumber: newNumber
      });
    }

    if (paragraph.manual || (paragraph.native && oldNumber !== newNumber)) {
      changes.push({
        paragraphIndex: paragraph.index,
        level,
        action: paragraph.manual ? "convert-manual" : "renumber-native",
        oldNumber,
        newNumber,
        textPreview: preview(paragraph),
        confidence: anomalyCode ? "review" : "high",
        bookmarkName,
        semanticKey: key
      });
    }
    void proposedOrdinal;
  }
  return { changes, targets, anomalies };
}
