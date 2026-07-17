const XML_ENTITIES: Record<string, string> = {
  amp: "&",
  apos: "'",
  gt: ">",
  lt: "<",
  quot: '"'
};

export function decodeXml(value: string): string {
  return value.replace(/&(#x[\da-f]+|#\d+|amp|apos|gt|lt|quot);/gi, (match, entity: string) => {
    if (entity.startsWith("#x")) return String.fromCodePoint(Number.parseInt(entity.slice(2), 16));
    if (entity.startsWith("#")) return String.fromCodePoint(Number.parseInt(entity.slice(1), 10));
    return XML_ENTITIES[entity.toLowerCase()] ?? match;
  });
}

export function encodeXml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

export function visibleText(xml: string): string {
  const pieces: string[] = [];
  const tokenPattern = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>|<w:br(?:\s[^>]*)?\s*\/>/g;
  for (const match of xml.matchAll(tokenPattern)) {
    if (match[1] !== undefined) pieces.push(decodeXml(match[1]));
    else if (match[0].startsWith("<w:tab")) pieces.push("\t");
    else pieces.push("\n");
  }
  return pieces.join("");
}

export function xmlAttribute(xml: string, qualifiedName: string): string | null {
  const escaped = qualifiedName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return xml.match(new RegExp(`${escaped}=["']([^"']+)["']`))?.[1] ?? null;
}

export function tagVal(xml: string, tagName: string): string | null {
  const escaped = tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const tag = xml.match(new RegExp(`<${escaped}\\b[^>]*\\bw:val=["']([^"']+)["'][^>]*/?>`));
  return tag?.[1] ?? null;
}

export function childXml(xml: string, tagName: string): string | null {
  const escaped = tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return xml.match(new RegExp(`<${escaped}\\b[^>]*>[\\s\\S]*?<\\/${escaped}>`))?.[0] ?? null;
}

export function stripVisiblePrefix(paragraphXml: string, count: number): string {
  let remaining = count;
  return paragraphXml.replace(/<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>/g, (whole, attrs: string | undefined, inner: string | undefined) => {
    if (remaining <= 0) return whole;
    if (inner === undefined) {
      remaining -= 1;
      return remaining >= 0 ? "" : whole;
    }
    const decoded = decodeXml(inner);
    if (decoded.length <= remaining) {
      remaining -= decoded.length;
      return "";
    }
    const next = decoded.slice(remaining);
    remaining = 0;
    const nextAttrs = /xml:space=/.test(attrs ?? "") || /^\s|\s$/.test(next)
      ? attrs ?? ' xml:space="preserve"'
      : attrs ?? "";
    return `<w:t${nextAttrs}>${encodeXml(next)}</w:t>`;
  });
}

export function paragraphProperties(paragraphXml: string): string | null {
  return paragraphXml.match(/<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/)?.[0] ?? null;
}

export function setParagraphNumbering(
  paragraphXml: string,
  numId: number,
  level: number,
  styleId: string
): string {
  const numbering = `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>`;
  const properties = paragraphProperties(paragraphXml);
  if (!properties) {
    return paragraphXml.replace(/^(<w:p\b[^>]*>)/, `$1<w:pPr><w:pStyle w:val="${styleId}"/>${numbering}</w:pPr>`);
  }
  let next = properties.replace(/<w:numPr\b[^>]*>[\s\S]*?<\/w:numPr>/g, "");
  if (/<w:pStyle\b/.test(next)) {
    next = next.replace(/<w:pStyle\b[^>]*\/>/, `<w:pStyle w:val="${styleId}"/>`);
  } else {
    next = next.replace(/^<w:pPr\b[^>]*>/, `$&<w:pStyle w:val="${styleId}"/>`);
  }
  next = next.replace(/<\/w:pPr>$/, `${numbering}</w:pPr>`);
  return paragraphXml.replace(properties, next);
}

export function addBookmark(paragraphXml: string, id: number, name: string): string {
  if (paragraphXml.includes(`w:name="${name}"`)) return paragraphXml;
  const start = `<w:bookmarkStart w:id="${id}" w:name="${name}"/>`;
  const end = `<w:bookmarkEnd w:id="${id}"/>`;
  const properties = paragraphProperties(paragraphXml);
  if (properties) {
    return paragraphXml.replace(properties, `${properties}${start}`).replace(/<\/w:p>$/, `${end}</w:p>`);
  }
  return paragraphXml.replace(/^(<w:p\b[^>]*>)/, `$1${start}`).replace(/<\/w:p>$/, `${end}</w:p>`);
}

function textRun(value: string, runProperties: string): string {
  if (!value) return "";
  const space = /^\s|\s$/.test(value) ? ' xml:space="preserve"' : "";
  return `<w:r>${runProperties}<w:t${space}>${encodeXml(value)}</w:t></w:r>`;
}

export function replaceTextWithRefField(
  paragraphXml: string,
  display: string,
  bookmarkName: string
): { xml: string; replaced: boolean } {
  let replaced = false;
  const xml = paragraphXml.replace(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g, (runXml) => {
    if (replaced || /<w:(?:instrText|fldChar)\b/.test(runXml)) return runXml;
    const textMatches = [...runXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)];
    if (textMatches.length !== 1) return runXml;
    const textMatch = textMatches[0];
    if (!textMatch || textMatch.index === undefined) return runXml;
    const decoded = decodeXml(textMatch[1] ?? "");
    const index = decoded.indexOf(display);
    if (index < 0) return runXml;
    const runProperties = runXml.match(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/)?.[0] ?? "";
    const before = decoded.slice(0, index);
    const after = decoded.slice(index + display.length);
    const instruction = ` REF ${bookmarkName} \\r \\h `;
    replaced = true;
    return [
      textRun(before, runProperties),
      `<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>`,
      `<w:r><w:instrText xml:space="preserve">${encodeXml(instruction)}</w:instrText></w:r>`,
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`,
      textRun(display, runProperties),
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>`,
      textRun(after, runProperties)
    ].join("");
  });
  return { xml, replaced };
}

export function maxNumericAttribute(xml: string, qualifiedName: string): number {
  const escaped = qualifiedName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  let max = 0;
  for (const match of xml.matchAll(new RegExp(`${escaped}=["'](\\d+)["']`, "g"))) {
    max = Math.max(max, Number(match[1]));
  }
  return max;
}
