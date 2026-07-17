import JSZip from "jszip";
import { detectManualToken } from "./tokens.js";
import type { TextInvariant } from "./types.js";
import { visibleText } from "./xml.js";

export interface SyntheticParagraph {
  text: string;
  table?: boolean;
  tracked?: "insert" | "delete";
}

function escape(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function paragraphXml(paragraph: SyntheticParagraph): string {
  const run = `<w:r><w:t xml:space="preserve">${escape(paragraph.text)}</w:t></w:r>`;
  const content = paragraph.tracked === "insert"
    ? `<w:ins w:id="1" w:author="Fixture" w:date="2026-01-01T00:00:00Z">${run}</w:ins>`
    : paragraph.tracked === "delete"
      ? `<w:del w:id="1" w:author="Fixture" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>${escape(paragraph.text)}</w:delText></w:r></w:del>`
      : run;
  return `<w:p>${content}</w:p>`;
}

export async function createSyntheticDocx(paragraphs: SyntheticParagraph[]): Promise<Uint8Array> {
  const zip = new JSZip();
  const body = paragraphs.map((paragraph) => paragraph.table
    ? `<w:tbl><w:tr><w:tc>${paragraphXml(paragraph)}<w:tcPr/></w:tc></w:tr></w:tbl>`
    : paragraphXml(paragraph)).join("");
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr/></w:body></w:document>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`);
  zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>`);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

export async function extractParagraphText(bytes: Uint8Array, removeManualTokens = false): Promise<string[]> {
  const zip = await JSZip.loadAsync(bytes);
  const documentXml = await zip.file("word/document.xml")?.async("string");
  if (!documentXml) return [];
  return [...documentXml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].map((match) => {
    const text = visibleText(match[0]);
    const token = removeManualTokens ? detectManualToken(text) : null;
    return token ? text.slice(token.stripLength) : text;
  });
}

export async function compareTextModuloNumberTokens(before: Uint8Array, after: Uint8Array): Promise<TextInvariant> {
  const beforeText = await extractParagraphText(before, true);
  const afterText = await extractParagraphText(after, true);
  return { before: beforeText, after: afterText, equal: JSON.stringify(beforeText) === JSON.stringify(afterText) };
}

export async function validateDocxPackage(bytes: Uint8Array): Promise<string[]> {
  const errors: string[] = [];
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
  } catch {
    return ["Package is not a readable ZIP archive."];
  }
  const required = ["[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/_rels/document.xml.rels"];
  for (const path of required) if (!zip.file(path)) errors.push(`Missing ${path}.`);
  const document = await zip.file("word/document.xml")?.async("string");
  if (document && (!document.includes("<w:document") || !document.includes("</w:document>"))) errors.push("word/document.xml is malformed.");
  const numbering = await zip.file("word/numbering.xml")?.async("string");
  if (numbering && (!numbering.includes("<w:numbering") || !numbering.includes("</w:numbering>"))) errors.push("word/numbering.xml is malformed.");
  return errors;
}
