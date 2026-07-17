import JSZip from "jszip";
import { LegalDownError } from "./errors.js";
import { decodeXml, encodeXml } from "./xml.js";

const PACKAGE_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return btoa(binary);
}

function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64.replace(/\s+/g, ""));
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}

function contentTypeMap(xml: string): { defaults: Map<string, string>; overrides: Map<string, string> } {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const match of xml.matchAll(/<Default\b[^>]*\bExtension=["']([^"']+)["'][^>]*\bContentType=["']([^"']+)["'][^>]*\/>/g)) {
    defaults.set((match[1] ?? "").toLowerCase(), match[2] ?? "application/octet-stream");
  }
  for (const match of xml.matchAll(/<Override\b[^>]*\bPartName=["']([^"']+)["'][^>]*\bContentType=["']([^"']+)["'][^>]*\/>/g)) {
    overrides.set(match[1] ?? "", match[2] ?? "application/octet-stream");
  }
  return { defaults, overrides };
}

export async function flatOpcToDocx(flatOpc: string): Promise<Uint8Array> {
  if (!/<pkg:package\b/.test(flatOpc)) {
    throw new LegalDownError("UNSUPPORTED_DOCUMENT", "Word did not return a Flat OPC document package.");
  }
  const zip = new JSZip();
  for (const partMatch of flatOpc.matchAll(/<pkg:part\b([^>]*)>([\s\S]*?)<\/pkg:part>/g)) {
    const attributes = partMatch[1] ?? "";
    const body = partMatch[2] ?? "";
    const encodedName = attributes.match(/pkg:name=["']([^"']+)["']/)?.[1];
    if (!encodedName) continue;
    const name = decodeXml(encodedName).replace(/^\//, "");
    const xmlData = body.match(/<pkg:xmlData>([\s\S]*?)<\/pkg:xmlData>/)?.[1];
    const binaryData = body.match(/<pkg:binaryData>([\s\S]*?)<\/pkg:binaryData>/)?.[1];
    if (xmlData !== undefined) {
      zip.file(name, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xmlData.trim()}`);
    } else if (binaryData !== undefined) {
      zip.file(name, base64ToBytes(binaryData));
    }
  }
  if (!zip.file("word/document.xml")) {
    throw new LegalDownError("UNSUPPORTED_DOCUMENT", "The Word OOXML package does not contain word/document.xml.");
  }
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

export async function docxToFlatOpc(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  const contentTypesXml = await zip.file("[Content_Types].xml")?.async("string");
  if (!contentTypesXml) throw new LegalDownError("INVALID_DOCX", "The repaired package has no content-type manifest.");
  const types = contentTypeMap(contentTypesXml);
  const parts: string[] = [];
  for (const [name, entry] of Object.entries(zip.files).sort(([left], [right]) => left.localeCompare(right))) {
    if (entry.dir) continue;
    const partName = `/${name}`;
    const extension = name.includes(".") ? name.split(".").at(-1)?.toLowerCase() ?? "" : "";
    const contentType = types.overrides.get(partName) ?? types.defaults.get(extension) ?? "application/octet-stream";
    const attributes = `pkg:name="${encodeXml(partName)}" pkg:contentType="${encodeXml(contentType)}"`;
    if (extension === "xml" || extension === "rels") {
      const xml = await entry.async("string");
      const withoutDeclaration = xml.replace(/^\s*<\?xml[^?]*\?>\s*/, "");
      parts.push(`<pkg:part ${attributes}><pkg:xmlData>${withoutDeclaration}</pkg:xmlData></pkg:part>`);
    } else {
      const data = await entry.async("uint8array");
      parts.push(`<pkg:part ${attributes}><pkg:binaryData>${bytesToBase64(data)}</pkg:binaryData></pkg:part>`);
    }
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pkg:package xmlns:pkg="${PACKAGE_NS}">${parts.join("")}</pkg:package>`;
}
