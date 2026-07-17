import type { ParagraphInventory } from "./types.js";
import { NumberingSimulator } from "./numbering.js";
import { detectManualToken, resolveAmbiguousTokens } from "./tokens.js";
import { paragraphProperties, tagVal, visibleText } from "./xml.js";

function integerVal(xml: string, tag: string, attribute: string): number | null {
  const escaped = tag.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const escapedAttribute = attribute.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const value = xml.match(new RegExp(`<${escaped}\\b[^>]*${escapedAttribute}=["'](-?\\d+)["'][^>]*/?>`))?.[1];
  return value === undefined ? null : Number(value);
}

function isInsideTable(documentXml: string, offset: number): boolean {
  const before = documentXml.slice(0, offset);
  return before.lastIndexOf("<w:tbl") > before.lastIndexOf("</w:tbl>");
}

export function inventoryDocument(documentXml: string, numberingXml: string | null): ParagraphInventory[] {
  const simulator = new NumberingSimulator(numberingXml);
  const paragraphs: ParagraphInventory[] = [];
  for (const match of documentXml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)) {
    const xml = match[0];
    const properties = paragraphProperties(xml) ?? "";
    const numPr = properties.match(/<w:numPr\b[^>]*>[\s\S]*?<\/w:numPr>/)?.[0] ?? "";
    const numId = integerVal(numPr, "w:numId", "w:val");
    const level = integerVal(numPr, "w:ilvl", "w:val") ?? 0;
    const native = numId === null ? null : simulator.next(numId, level);
    const text = visibleText(xml);
    paragraphs.push({
      index: paragraphs.length,
      text,
      styleId: tagVal(properties, "w:pStyle"),
      indentLeft: integerVal(properties, "w:ind", "w:left") ?? integerVal(properties, "w:ind", "w:start"),
      indentHanging: integerVal(properties, "w:ind", "w:hanging"),
      inTable: isInsideTable(documentXml, match.index ?? 0),
      hasTrackedChanges: /<w:(?:ins|del)\b/.test(xml),
      hasFields: /<w:(?:instrText|fldChar|fldSimple)\b/.test(xml),
      hasRefField: /<w:instrText\b[^>]*>[\s\S]*?\bREF\s+/i.test(xml) || /<w:fldSimple\b[^>]*\bw:instr=["'][^"']*\bREF\s+/i.test(xml),
      manual: numId === null ? detectManualToken(text) : null,
      native: native && numId !== null ? {
        numId,
        level,
        rendered: native.rendered,
        counters: native.counters,
        format: native.format
      } : null
    });
  }
  resolveAmbiguousTokens(paragraphs.map((paragraph) => paragraph.manual), paragraphs.map((paragraph) => paragraph.indentLeft));
  return paragraphs;
}
