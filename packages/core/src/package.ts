import JSZip from "jszip";
import { WordOrderError } from "./errors.js";

export interface DocxPackage {
  zip: JSZip;
  original: Uint8Array;
  documentXml: string;
  numberingXml: string | null;
  stylesXml: string | null;
  settingsXml: string | null;
  relationshipsXml: string | null;
  contentTypesXml: string;
  hasTrackedChanges: boolean;
}

async function readRequired(zip: JSZip, path: string): Promise<string> {
  const file = zip.file(path);
  if (!file) throw new WordOrderError("INVALID_DOCX", `The document is missing the required OOXML part ${path}.`);
  return file.async("string");
}

async function readOptional(zip: JSZip, path: string): Promise<string | null> {
  return zip.file(path)?.async("string") ?? null;
}

export async function loadDocx(bytes: Uint8Array): Promise<DocxPackage> {
  if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    throw new WordOrderError("INVALID_DOCX", "This file is not a valid .docx package.");
  }
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
  } catch {
    throw new WordOrderError("INVALID_DOCX", "The .docx ZIP package is corrupt or unreadable.");
  }
  const documentXml = await readRequired(zip, "word/document.xml");
  const contentTypesXml = await readRequired(zip, "[Content_Types].xml");
  const wordXmlFiles = Object.keys(zip.files).filter((name) => name.startsWith("word/") && name.endsWith(".xml"));
  let hasTrackedChanges = false;
  for (const name of wordXmlFiles) {
    const xml = await readOptional(zip, name);
    if (xml && /<w:(?:ins|del)\b/.test(xml)) {
      hasTrackedChanges = true;
      break;
    }
  }
  return {
    zip,
    original: bytes,
    documentXml,
    numberingXml: await readOptional(zip, "word/numbering.xml"),
    stylesXml: await readOptional(zip, "word/styles.xml"),
    settingsXml: await readOptional(zip, "word/settings.xml"),
    relationshipsXml: await readOptional(zip, "word/_rels/document.xml.rels"),
    contentTypesXml,
    hasTrackedChanges
  };
}

export function fingerprint(bytes: Uint8Array): string {
  let first = 0x811c9dc5;
  let second = 0x9e3779b9;
  for (const byte of bytes) {
    first = Math.imul(first ^ byte, 0x01000193) >>> 0;
    second = Math.imul(second ^ (byte + 1), 0x85ebca6b) >>> 0;
  }
  return `${first.toString(16).padStart(8, "0")}${second.toString(16).padStart(8, "0")}`;
}
