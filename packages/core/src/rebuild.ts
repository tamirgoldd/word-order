import type { DocxPackage } from "./package.js";
import type { RepairPlan } from "./types.js";
import { normalizeMargins, normalizeParagraphFormatting } from "./formatting.js";
import {
  addBookmark,
  maxNumericAttribute,
  replaceTextWithRefField,
  setParagraphNumbering,
  stripVisiblePrefix
} from "./xml.js";

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

function levelXml(level: number, start: number, format: string, text: string, indent: number, legal = false): string {
  const hanging = level < 2 ? 0 : 360;
  return `<w:lvl w:ilvl="${level}"><w:start w:val="${start}"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>${legal ? "<w:isLgl/>" : ""}<w:pPr><w:tabs><w:tab w:val="num" w:pos="${indent}"/></w:tabs><w:ind w:left="${indent}" w:hanging="${hanging}"/></w:pPr></w:lvl>`;
}

function addNumberingDefinition(existing: string | null, plan: RepairPlan): { xml: string; abstractId: number; numId: number } {
  const base = existing ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${WORD_NS}"></w:numbering>`;
  const abstractId = maxNumericAttribute(base, "w:abstractNumId") + 1;
  const numId = maxNumericAttribute(base, "w:numId") + 1;
  const articleText = `${plan.profile.articleLabel} %1`;
  const topLevelFormat = plan.profile.topLevelKind === "flat-section" ? "decimal" : plan.profile.articleFormat;
  const topLevelText = plan.profile.topLevelKind === "flat-section" ? "%1." : articleText;
  const sectionText = plan.profile.sectionLabel === "none" ? "%1.%2" : `${plan.profile.sectionLabel} %1.%2`;
  const abstract = [
    `<w:abstractNum w:abstractNumId="${abstractId}">`,
    `<w:nsid w:val="4C444F57"/><w:multiLevelType w:val="multilevel"/><w:name w:val="Legal Down"/>`,
    levelXml(0, 1, topLevelFormat, topLevelText, plan.profile.levelIndents[0] ?? 0),
    levelXml(1, 1, "decimalZero", sectionText, plan.profile.levelIndents[1] ?? 720, true),
    levelXml(2, 1, "lowerLetter", "(%3)", plan.profile.levelIndents[2] ?? 1440),
    levelXml(3, 1, "lowerRoman", "(%4)", plan.profile.levelIndents[3] ?? 2160),
    `</w:abstractNum>`,
    `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractId}"/></w:num>`
  ].join("");
  return { xml: base.replace(/<\/w:numbering>\s*$/, `${abstract}</w:numbering>`), abstractId, numId };
}

function ensureStyles(existing: string | null, plan: RepairPlan): string {
  let xml = existing ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${WORD_NS}"></w:styles>`;
  const font = plan.formattingProfile.fontFamily;
  const styles = [
    { id: "LDArticle", name: "Legal Down Article", pPr: `<w:keepNext/><w:spacing w:before="240" w:after="120"/><w:ind w:left="${plan.profile.levelIndents[0] ?? 0}"/>`, rPr: `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:b/><w:sz w:val="${plan.formattingProfile.headingSizeHalfPoints}"/><w:szCs w:val="${plan.formattingProfile.headingSizeHalfPoints}"/>` },
    { id: "LDSection", name: "Legal Down Section", pPr: `<w:keepNext/><w:spacing w:before="240" w:after="120"/><w:ind w:left="${plan.profile.topLevelKind === "flat-section" ? 0 : plan.profile.levelIndents[1] ?? 720}"/>`, rPr: `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:b/><w:sz w:val="${plan.formattingProfile.headingSizeHalfPoints}"/><w:szCs w:val="${plan.formattingProfile.headingSizeHalfPoints}"/>` },
    { id: "LDClause", name: "Legal Down Clause", pPr: `<w:spacing w:after="80"/><w:ind w:left="${plan.profile.levelIndents[2] ?? 1440}"/>`, rPr: `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="${plan.formattingProfile.bodySizeHalfPoints}"/><w:szCs w:val="${plan.formattingProfile.bodySizeHalfPoints}"/>` },
    { id: "LDSubclause", name: "Legal Down Subclause", pPr: `<w:spacing w:after="80"/><w:ind w:left="${plan.profile.levelIndents[3] ?? 2160}"/>`, rPr: `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="${plan.formattingProfile.bodySizeHalfPoints}"/><w:szCs w:val="${plan.formattingProfile.bodySizeHalfPoints}"/>` },
    { id: "LDTitle", name: "Legal Down Title", pPr: `<w:jc w:val="center"/><w:keepNext/><w:spacing w:after="80"/>`, rPr: `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:b/><w:sz w:val="${plan.formattingProfile.titleSizeHalfPoints}"/><w:szCs w:val="${plan.formattingProfile.titleSizeHalfPoints}"/>` },
    { id: "LDSubtitle", name: "Legal Down Subtitle", pPr: `<w:jc w:val="center"/><w:keepNext/><w:spacing w:after="240"/>`, rPr: `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:i/><w:sz w:val="${plan.formattingProfile.subtitleSizeHalfPoints}"/><w:szCs w:val="${plan.formattingProfile.subtitleSizeHalfPoints}"/>` },
    { id: "LDBody", name: "Legal Down Body", pPr: `<w:jc w:val="both"/><w:spacing w:after="120" w:line="276" w:lineRule="auto"/>`, rPr: `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="${plan.formattingProfile.bodySizeHalfPoints}"/><w:szCs w:val="${plan.formattingProfile.bodySizeHalfPoints}"/>` },
    { id: "LDSignature", name: "Legal Down Signature", pPr: `<w:jc w:val="left"/><w:spacing w:before="80" w:after="120"/>`, rPr: `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="${plan.formattingProfile.bodySizeHalfPoints}"/><w:szCs w:val="${plan.formattingProfile.bodySizeHalfPoints}"/>` }
  ];
  const additions = styles.flatMap((style) => {
    const styleId = style.id;
    if (new RegExp(`w:styleId=["']${styleId}["']`).test(xml)) return [];
    return [`<w:style w:type="paragraph" w:customStyle="1" w:styleId="${styleId}"><w:name w:val="${style.name}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr>${style.pPr}</w:pPr><w:rPr>${style.rPr}</w:rPr></w:style>`];
  }).join("");
  if (additions) xml = xml.replace(/<\/w:styles>\s*$/, `${additions}</w:styles>`);
  return xml;
}

function ensureSettings(existing: string | null): string {
  const base = existing ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${WORD_NS}"></w:settings>`;
  if (/<w:updateFields\b/.test(base)) {
    return base.replace(/<w:updateFields\b[^>]*\/>/, '<w:updateFields w:val="true"/>');
  }
  return base.replace(/<\/w:settings>\s*$/, '<w:updateFields w:val="true"/></w:settings>');
}

interface PartRelationship {
  target: string;
  type: string;
}

function ensureRelationships(existing: string | null, needed: PartRelationship[]): string {
  let xml = existing ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL_NS}"></Relationships>`;
  let nextId = 1;
  for (const match of xml.matchAll(/Id=["']rId(\d+)["']/g)) nextId = Math.max(nextId, Number(match[1]) + 1);
  for (const relationship of needed) {
    if (new RegExp(`Target=["']${relationship.target.replace(".", "\\.")}["']`).test(xml)) continue;
    const addition = `<Relationship Id="rId${nextId}" Type="${relationship.type}" Target="${relationship.target}"/>`;
    nextId += 1;
    xml = xml.replace(/<\/Relationships>\s*$/, `${addition}</Relationships>`);
  }
  return xml;
}

function ensureContentTypes(existing: string, parts: Array<{ name: string; type: string }>): string {
  let xml = existing;
  for (const part of parts) {
    if (xml.includes(`PartName="${part.name}"`) || xml.includes(`PartName='${part.name}'`)) continue;
    xml = xml.replace(/<\/Types>\s*$/, `<Override PartName="${part.name}" ContentType="${part.type}"/></Types>`);
  }
  return xml;
}

function styleForLevel(level: number, plan: RepairPlan): string {
  if (level === 0 && plan.profile.topLevelKind === "flat-section") return "LDSection";
  return ["LDArticle", "LDSection", "LDClause", "LDSubclause"][Math.min(level, 3)] ?? "LDClause";
}

export async function rebuildDocx(docx: DocxPackage, plan: RepairPlan): Promise<Uint8Array> {
  const rebuildNumbering = plan.numbering.length > 0;
  const resolvedReferences = plan.crossReferences.filter((reference) => reference.status === "resolved");
  const rebuildFormatting = plan.formatting.length > 0;
  if (!rebuildNumbering && resolvedReferences.length === 0 && !rebuildFormatting) return docx.original.slice();

  let numberingXml = docx.numberingXml;
  let numId = -1;
  if (rebuildNumbering) {
    const numbering = addNumberingDefinition(numberingXml, plan);
    numberingXml = numbering.xml;
    numId = numbering.numId;
  }

  const changes = new Map(plan.numbering.map((change) => [change.paragraphIndex, change]));
  const targets = new Map(plan.targets.map((target) => [target.paragraphIndex, target]));
  const referencedTargets = new Set(resolvedReferences.flatMap((reference) => reference.targetParagraphIndex === null ? [] : [reference.targetParagraphIndex]));
  const referencesByParagraph = new Map<number, typeof resolvedReferences>();
  for (const reference of resolvedReferences) {
    const list = referencesByParagraph.get(reference.paragraphIndex) ?? [];
    list.push(reference);
    referencesByParagraph.set(reference.paragraphIndex, list);
  }
  const formattingRoles = new Map(plan.formatting.flatMap((change) => change.paragraphIndex === null || change.role === "section"
    ? []
    : [[change.paragraphIndex, change.role] as const]));
  let bookmarkId = maxNumericAttribute(docx.documentXml, "w:id") + 1;
  let paragraphIndex = 0;
  const documentXml = docx.documentXml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (originalParagraph) => {
    const index = paragraphIndex;
    paragraphIndex += 1;
    let paragraph = originalParagraph;
    const change = changes.get(index);
    const target = targets.get(index);
    if (change?.action === "convert-manual") {
      const stripLength = plan.inventory[index]?.manual?.stripLength ?? 0;
      paragraph = stripVisiblePrefix(paragraph, stripLength);
    }
    if (rebuildNumbering && target) {
      paragraph = setParagraphNumbering(paragraph, numId, target.level, styleForLevel(target.level, plan));
    }
    if (target && referencedTargets.has(index)) {
      paragraph = addBookmark(paragraph, bookmarkId, target.bookmarkName);
      bookmarkId += 1;
    }
    for (const reference of referencesByParagraph.get(index) ?? []) {
      if (!reference.bookmarkName) continue;
      paragraph = replaceTextWithRefField(paragraph, reference.display, reference.bookmarkName).xml;
    }
    const formattingRole = formattingRoles.get(index);
    if (formattingRole) paragraph = normalizeParagraphFormatting(paragraph, formattingRole);
    return paragraph;
  });

  const normalizedDocumentXml = rebuildFormatting && plan.formatting.some((change) => change.category === "margins")
    ? normalizeMargins(documentXml, plan.formattingProfile)
    : documentXml;

  docx.zip.file("word/document.xml", normalizedDocumentXml);
  if (numberingXml) docx.zip.file("word/numbering.xml", numberingXml);
  docx.zip.file("word/styles.xml", ensureStyles(docx.stylesXml, plan));
  docx.zip.file("word/settings.xml", ensureSettings(docx.settingsXml));
  docx.zip.file("word/_rels/document.xml.rels", ensureRelationships(docx.relationshipsXml, [
    { target: "numbering.xml", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" },
    { target: "styles.xml", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" },
    { target: "settings.xml", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" }
  ]));
  docx.zip.file("[Content_Types].xml", ensureContentTypes(docx.contentTypesXml, [
    { name: "/word/numbering.xml", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml" },
    { name: "/word/styles.xml", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" },
    { name: "/word/settings.xml", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml" }
  ]));
  return docx.zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
}
