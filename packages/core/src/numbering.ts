import { childXml, tagVal, xmlAttribute } from "./xml.js";
import { numberToLetters, numberToRoman } from "./tokens.js";

interface LevelDefinition {
  level: number;
  start: number;
  format: string;
  text: string;
  isLegal: boolean;
}

interface NumberInstance {
  abstractId: number;
  starts: Map<number, number>;
}

interface ParsedNumbering {
  abstracts: Map<number, Map<number, LevelDefinition>>;
  instances: Map<number, NumberInstance>;
}

function parseLevels(abstractXml: string): Map<number, LevelDefinition> {
  const levels = new Map<number, LevelDefinition>();
  for (const match of abstractXml.matchAll(/<w:lvl\b[^>]*\bw:ilvl=["'](\d+)["'][^>]*>[\s\S]*?<\/w:lvl>/g)) {
    const level = Number(match[1]);
    const xml = match[0];
    levels.set(level, {
      level,
      start: Number(tagVal(xml, "w:start") ?? 1),
      format: tagVal(xml, "w:numFmt") ?? "decimal",
      text: tagVal(xml, "w:lvlText") ?? `%${level + 1}`,
      isLegal: /<w:isLgl\b/.test(xml)
    });
  }
  return levels;
}

export function parseNumbering(xml: string | null): ParsedNumbering {
  const abstracts = new Map<number, Map<number, LevelDefinition>>();
  const instances = new Map<number, NumberInstance>();
  if (!xml) return { abstracts, instances };

  for (const match of xml.matchAll(/<w:abstractNum\b[^>]*\bw:abstractNumId=["'](\d+)["'][^>]*>[\s\S]*?<\/w:abstractNum>/g)) {
    abstracts.set(Number(match[1]), parseLevels(match[0]));
  }
  for (const match of xml.matchAll(/<w:num\b[^>]*\bw:numId=["'](\d+)["'][^>]*>[\s\S]*?<\/w:num>/g)) {
    const xmlFragment = match[0];
    const abstractTag = xmlFragment.match(/<w:abstractNumId\b[^>]*\/>/)?.[0] ?? "";
    const abstractId = Number(xmlAttribute(abstractTag, "w:val") ?? -1);
    const starts = new Map<number, number>();
    for (const override of xmlFragment.matchAll(/<w:lvlOverride\b[^>]*\bw:ilvl=["'](\d+)["'][^>]*>[\s\S]*?<\/w:lvlOverride>/g)) {
      const startXml = childXml(override[0], "w:startOverride") ?? override[0].match(/<w:startOverride\b[^>]*\/>/)?.[0] ?? "";
      const value = xmlAttribute(startXml, "w:val");
      if (value !== null) starts.set(Number(override[1]), Number(value));
    }
    instances.set(Number(match[1]), { abstractId, starts });
  }
  return { abstracts, instances };
}

function formatCounter(value: number, format: string): string {
  switch (format) {
    case "decimalZero": return String(value).padStart(2, "0");
    case "upperRoman": return numberToRoman(value);
    case "lowerRoman": return numberToRoman(value).toLowerCase();
    case "upperLetter": return numberToLetters(value, true);
    case "lowerLetter": return numberToLetters(value);
    default: return String(value);
  }
}

export class NumberingSimulator {
  readonly #parsed: ParsedNumbering;
  readonly #counters = new Map<number, number[]>();

  constructor(numberingXml: string | null) {
    this.#parsed = parseNumbering(numberingXml);
  }

  next(numId: number, level: number): { rendered: string; counters: number[]; format: string } | null {
    const instance = this.#parsed.instances.get(numId);
    if (!instance) return null;
    const levels = this.#parsed.abstracts.get(instance.abstractId);
    const definition = levels?.get(level);
    if (!levels || !definition) return null;
    const counters = this.#counters.get(numId) ?? [];
    const start = instance.starts.get(level) ?? definition.start;
    counters[level] = counters[level] === undefined ? start : (counters[level] ?? start - 1) + 1;
    counters.length = level + 1;
    for (let parent = 0; parent < level; parent += 1) {
      if (counters[parent] === undefined) counters[parent] = levels.get(parent)?.start ?? 1;
    }
    this.#counters.set(numId, counters);
    const rendered = definition.text.replace(/%(\d+)/g, (_whole, oneBased: string) => {
      const referencedLevel = Number(oneBased) - 1;
      const referenced = levels.get(referencedLevel);
      const format = definition.isLegal && referencedLevel < level ? "decimal" : referenced?.format ?? "decimal";
      return formatCounter(counters[referencedLevel] ?? 0, format);
    });
    return { rendered, counters: [...counters], format: definition.format };
  }
}
