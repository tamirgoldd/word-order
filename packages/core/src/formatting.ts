import type {
  FormattingCategory,
  FormattingChange,
  FormattingProfile,
  FormattingRole,
  NumberedTarget,
  ParagraphInventory,
  PlanWarning
} from "./types.js";
import { paragraphProperties, tagVal, visibleText } from "./xml.js";

export const LEGAL_FORMATTING_PROFILE: FormattingProfile = {
  fontFamily: "Times New Roman",
  titleSizeHalfPoints: 36,
  subtitleSizeHalfPoints: 20,
  headingSizeHalfPoints: 24,
  bodySizeHalfPoints: 22,
  marginTwips: 1440
};

interface FormattingAnalysis {
  changes: FormattingChange[];
  warnings: PlanWarning[];
}

function unique(values: Array<string | null>): string[] {
  return [...new Set(values.filter((value): value is string => Boolean(value)))];
}

function attribute(xml: string, name: string): string | null {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return xml.match(new RegExp(`${escaped}=["']([^"']+)["']`))?.[1] ?? null;
}

function directRunProperties(paragraphXml: string): string[] {
  return [...paragraphXml.matchAll(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/g)].map((match) => match[0]);
}

function valuesForTag(properties: string[], tag: string, attributeName = "w:val"): string[] {
  return unique(properties.flatMap((property) => [...property.matchAll(new RegExp(`<w:${tag}\\b[^>]*>`, "g"))]
    .map((match) => attribute(match[0], attributeName))));
}

function fonts(properties: string[]): string[] {
  return unique(properties.flatMap((property) => [...property.matchAll(/<w:rFonts\b[^>]*\/?\s*>/g)].flatMap((match) => [
    attribute(match[0], "w:ascii"),
    attribute(match[0], "w:hAnsi"),
    attribute(match[0], "w:eastAsia"),
    attribute(match[0], "w:cs")
  ])));
}

function headingLike(paragraph: ParagraphInventory): boolean {
  if (paragraph.manual?.kind !== "decimal-item" || paragraph.manual.suffix !== ".") return false;
  const text = paragraph.text.slice(paragraph.manual.stripLength).trim();
  const words = text.split(/\s+/).filter(Boolean);
  return text.length > 0 && text.length <= 100 && words.length <= 12 && !/[;:]/.test(text);
}

export function isFlatSectionHeading(paragraph: ParagraphInventory): boolean {
  return headingLike(paragraph);
}

function classifyRoles(inventory: ParagraphInventory[], targets: NumberedTarget[]): Map<number, FormattingRole> {
  const roles = new Map<number, FormattingRole>();
  const nonempty = inventory.filter((paragraph) => paragraph.text.trim() && !paragraph.inTable);
  const title = nonempty[0];
  if (title) {
    const titleText = title.text.trim();
    const titleWords = titleText.split(/\s+/).filter(Boolean);
    if (titleText.length <= 120 && titleWords.length <= 12 && !/[.!?;:]$/.test(titleText)) roles.set(title.index, "title");
  }
  const subtitle = nonempty[1];
  if (title && roles.get(title.index) === "title" && subtitle && /^\(\s*the\s+["“]?agreement/i.test(subtitle.text.trim())) roles.set(subtitle.index, "subtitle");
  const headingIndexes = new Set(targets.map((target) => target.paragraphIndex));
  for (const paragraph of nonempty) {
    const text = paragraph.text.trim();
    if (headingIndexes.has(paragraph.index) || headingLike(paragraph)) roles.set(paragraph.index, "heading");
    else if (/^(?:IN WITNESS WHEREOF\b|(?:Company|Provider|Date)\s*:)/i.test(text)) roles.set(paragraph.index, "signature");
    else if (!roles.has(paragraph.index)) roles.set(paragraph.index, "body");
  }
  return roles;
}

function desiredStyle(role: FormattingRole): string {
  return {
    title: "LDTitle",
    subtitle: "LDSubtitle",
    heading: "LDSection",
    body: "LDBody",
    signature: "LDSignature"
  }[role];
}

function desiredAlignment(role: FormattingRole): string {
  if (role === "title" || role === "subtitle") return "center";
  if (role === "body") return "both";
  return "left";
}

function desiredSize(role: FormattingRole, profile: FormattingProfile): number {
  if (role === "title") return profile.titleSizeHalfPoints;
  if (role === "subtitle") return profile.subtitleSizeHalfPoints;
  if (role === "heading") return profile.headingSizeHalfPoints;
  return profile.bodySizeHalfPoints;
}

function addChange(
  changes: FormattingChange[],
  paragraph: ParagraphInventory,
  role: FormattingRole,
  category: FormattingCategory,
  oldValue: string,
  newValue: string
): void {
  changes.push({
    paragraphIndex: paragraph.index,
    role,
    category,
    oldValue,
    newValue,
    textPreview: paragraph.text.trim().slice(0, 100),
    confidence: "high"
  });
}

function marginValues(documentXml: string): number[][] {
  return [...documentXml.matchAll(/<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g)].map((section) => {
    const margin = section[0].match(/<w:pgMar\b[^>]*\/?\s*>/)?.[0] ?? "";
    return ["w:top", "w:right", "w:bottom", "w:left"].map((name) => Number(attribute(margin, name) ?? 1440));
  });
}

function shouldNormalize(documentXml: string): boolean {
  const runProperties = directRunProperties(documentXml);
  const fontFamilies = fonts(runProperties);
  const sizes = valuesForTag(runProperties, "sz");
  const alignments = unique([...documentXml.matchAll(/<w:jc\b[^>]*\/?\s*>/g)].map((match) => attribute(match[0], "w:val")));
  const indents = [...documentXml.matchAll(/<w:ind\b[^>]*\/?\s*>/g)].flatMap((match) => {
    const value = attribute(match[0], "w:left") ?? attribute(match[0], "w:start");
    return value === null ? [] : [Math.abs(Number(value))];
  });
  const margins = marginValues(documentXml).flat();
  const strangeFont = fontFamilies.some((font) => /comic sans|papyrus|curlz|jokerman|chiller/i.test(font));
  const inconsistentMargins = margins.some((margin) => margin < 720 || margin > 1800) || (margins.length > 0 && Math.max(...margins) - Math.min(...margins) > 480);
  return strangeFont
    || fontFamilies.length >= 3
    || sizes.length >= 4
    || /<w:highlight\b/.test(documentXml)
    || indents.some((indent) => indent > 720)
    || inconsistentMargins
    || (fontFamilies.length >= 2 && sizes.length >= 3 && alignments.length >= 3);
}

function hasUnintendedEmphasis(paragraphXml: string, role: FormattingRole): boolean {
  if (role === "title" || role === "subtitle" || role === "heading") return /<w:(?:b|i|u)\b/.test(paragraphXml);
  for (const match of paragraphXml.matchAll(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g)) {
    const run = match[0];
    const text = visibleText(run).trim();
    const intentionalBold = /^\[[^\]]+\]$/.test(text) || (/[A-Z]/.test(text) && !/[a-z]/.test(text) && text.replace(/[^A-Z]/g, "").length >= 3);
    if (/<w:u\b|<w:i\b/.test(run) || (/<w:b\b/.test(run) && !intentionalBold)) return true;
  }
  return false;
}

function longParagraphWarnings(inventory: ParagraphInventory[]): PlanWarning[] {
  return inventory.flatMap((paragraph) => {
    if (paragraph.inTable || paragraph.manual || paragraph.native) return [];
    const text = paragraph.text.trim();
    const wordCount = text.split(/\s+/).filter(Boolean).length;
    const sentenceBreaks = (text.match(/[.!?](?:["”')\]]+)?(?:\s|$)/g) ?? []).length;
    if (wordCount < 55 || sentenceBreaks > 1) return [];
    return [{
      code: "long-paragraph" as const,
      paragraphIndex: paragraph.index,
      message: `This ${wordCount}-word paragraph has no internal sentence break. Word Order preserved the wording and flagged it for editorial review.`
    }];
  });
}

export function analyzeFormatting(
  documentXml: string,
  inventory: ParagraphInventory[],
  targets: NumberedTarget[],
  profile: FormattingProfile = LEGAL_FORMATTING_PROFILE
): FormattingAnalysis {
  const warnings = longParagraphWarnings(inventory);
  if (!shouldNormalize(documentXml)) return { changes: [], warnings };
  const roles = classifyRoles(inventory, targets);
  const paragraphXml = [...documentXml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].map((match) => match[0]);
  const changes: FormattingChange[] = [];

  for (const paragraph of inventory) {
    const role = roles.get(paragraph.index);
    const xml = paragraphXml[paragraph.index] ?? "";
    if (!role || !xml) continue;
    const properties = paragraphProperties(xml) ?? "";
    const runProperties = directRunProperties(xml);
    const paragraphFonts = fonts(runProperties);
    const sizes = valuesForTag(runProperties, "sz");
    const alignment = tagVal(properties, "w:jc");
    const style = tagVal(properties, "w:pStyle");

    if (style !== desiredStyle(role)) addChange(changes, paragraph, role, "style", style ?? "direct formatting", desiredStyle(role));
    if (paragraphFonts.length) addChange(changes, paragraph, role, "font-family", paragraphFonts.join(", "), profile.fontFamily);
    if (sizes.length) addChange(changes, paragraph, role, "font-size", sizes.map((size) => `${Number(size) / 2} pt`).join(", "), `${desiredSize(role, profile) / 2} pt`);
    if (hasUnintendedEmphasis(xml, role)) addChange(changes, paragraph, role, "emphasis", "mixed bold, italic, or underline", role === "title" || role === "heading" ? "bold" : role === "subtitle" ? "italic" : "plain text with meaningful bold retained");
    if (alignment && alignment !== desiredAlignment(role)) addChange(changes, paragraph, role, "alignment", alignment, desiredAlignment(role));
    if (/<w:ind\b/.test(properties)) addChange(changes, paragraph, role, "indent", "direct paragraph indent", "style-defined indent");
    if (/<w:spacing\b/.test(properties)) addChange(changes, paragraph, role, "spacing", "direct paragraph spacing", "consistent legal-document spacing");
    if (/<w:highlight\b/.test(xml)) addChange(changes, paragraph, role, "highlight", "highlighted text", "no highlight");
  }

  const margins = marginValues(documentXml);
  for (const [sectionIndex, values] of margins.entries()) {
    if (values.every((value) => value === profile.marginTwips)) continue;
    changes.push({
      paragraphIndex: null,
      role: "section",
      category: "margins",
      oldValue: values.map((value) => `${(value / 1440).toFixed(2)} in`).join(" / "),
      newValue: "1.00 in on all sides",
      textPreview: `Section ${sectionIndex + 1}`,
      confidence: "high"
    });
  }
  return { changes, warnings };
}

function removeProperty(xml: string, tag: string): string {
  return xml.replace(new RegExp(`<w:${tag}\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/w:${tag}>)`, "g"), "");
}

function intentionalBold(runXml: string): boolean {
  const text = visibleText(runXml).trim();
  return /^\[[^\]]+\]$/.test(text) || (/[A-Z]/.test(text) && !/[a-z]/.test(text) && text.replace(/[^A-Z]/g, "").length >= 3);
}

function normalizeRuns(paragraphXml: string, role: FormattingRole): string {
  return paragraphXml.replace(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g, (run) => {
    if (/<w:(?:instrText|fldChar)\b/.test(run)) return run;
    const preserveBold = (role === "body" || role === "signature") && intentionalBold(run);
    const properties = run.match(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/)?.[0] ?? null;
    if (!properties) return preserveBold ? run.replace(/^(<w:r\b[^>]*>)/, "$1<w:rPr><w:b/></w:rPr>") : run;
    let next = properties;
    for (const tag of ["rFonts", "sz", "szCs", "b", "bCs", "i", "iCs", "u", "highlight", "color"]) next = removeProperty(next, tag);
    if (preserveBold) next = next.replace(/<\/w:rPr>$/, "<w:b/></w:rPr>");
    if (/^<w:rPr\b[^>]*>\s*<\/w:rPr>$/.test(next)) next = "";
    return run.replace(properties, next);
  });
}

function setStyle(paragraphXml: string, role: FormattingRole): string {
  const styleId = desiredStyle(role);
  const properties = paragraphProperties(paragraphXml);
  if (!properties) return paragraphXml.replace(/^(<w:p\b[^>]*>)/, `$1<w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>`);
  let next = properties;
  for (const tag of ["pStyle", "jc", "ind", "spacing", "tabs", "keepNext", "keepLines", "widowControl"]) next = removeProperty(next, tag);
  next = next.replace(/^<w:pPr\b[^>]*>/, `$&<w:pStyle w:val="${styleId}"/>`);
  return paragraphXml.replace(properties, next);
}

export function normalizeParagraphFormatting(paragraphXml: string, role: FormattingRole): string {
  return normalizeRuns(setStyle(paragraphXml, role), role);
}

function setAttribute(xml: string, name: string, value: number): string {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  if (new RegExp(`${escaped}=["'][^"']*["']`).test(xml)) return xml.replace(new RegExp(`${escaped}=["'][^"']*["']`), `${name}="${value}"`);
  return xml.replace(/\/?\s*>$/, (end) => ` ${name}="${value}"${end}`);
}

export function normalizeMargins(documentXml: string, profile: FormattingProfile): string {
  return documentXml.replace(/<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g, (section) => {
    const existing = section.match(/<w:pgMar\b[^>]*\/?\s*>/)?.[0];
    if (!existing) return section.replace(/<\/w:sectPr>$/, `<w:pgMar w:top="${profile.marginTwips}" w:right="${profile.marginTwips}" w:bottom="${profile.marginTwips}" w:left="${profile.marginTwips}"/></w:sectPr>`);
    let next = existing;
    for (const name of ["w:top", "w:right", "w:bottom", "w:left"]) next = setAttribute(next, name, profile.marginTwips);
    return section.replace(existing, next);
  });
}
