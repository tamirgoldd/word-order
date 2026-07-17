import type { ManualToken, TokenKind } from "./types.js";

const NUMBER_WORDS: Record<string, number> = {
  one: 1,
  two: 2,
  three: 3,
  four: 4,
  five: 5,
  six: 6,
  seven: 7,
  eight: 8,
  nine: 9,
  ten: 10,
  eleven: 11,
  twelve: 12,
  thirteen: 13,
  fourteen: 14,
  fifteen: 15,
  sixteen: 16,
  seventeen: 17,
  eighteen: 18,
  nineteen: 19,
  twenty: 20
};

export function romanToNumber(value: string): number {
  const roman = value.toUpperCase();
  if (!/^[IVXLCDM]+$/.test(roman)) return Number.NaN;
  const values: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  let total = 0;
  for (let index = 0; index < roman.length; index += 1) {
    const current = values[roman[index] ?? ""] ?? 0;
    const next = values[roman[index + 1] ?? ""] ?? 0;
    total += current < next ? -current : current;
  }
  return total;
}

export function numberToRoman(value: number): string {
  if (!Number.isInteger(value) || value <= 0 || value >= 4000) return String(value);
  const table: Array<[number, string]> = [
    [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"],
    [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]
  ];
  let rest = value;
  let result = "";
  for (const [amount, token] of table) {
    while (rest >= amount) {
      result += token;
      rest -= amount;
    }
  }
  return result;
}

export function lettersToNumber(value: string): number {
  let result = 0;
  for (const character of value.toLowerCase()) result = result * 26 + character.charCodeAt(0) - 96;
  return result;
}

export function numberToLetters(value: number, uppercase = false): string {
  let rest = value;
  let result = "";
  while (rest > 0) {
    rest -= 1;
    result = String.fromCharCode(97 + (rest % 26)) + result;
    rest = Math.floor(rest / 26);
  }
  return uppercase ? result.toUpperCase() : result;
}

function createToken(
  match: RegExpMatchArray,
  raw: string,
  kind: TokenKind,
  ordinal: number,
  components: number[],
  label: string,
  suffix = ""
): ManualToken {
  return {
    raw,
    stripLength: match[0].length,
    kind,
    ordinal,
    components,
    label,
    suffix
  };
}

export function detectManualToken(text: string): ManualToken | null {
  const article = text.match(/^\s*((ARTICLE|Article)\s+([A-Za-z]+|\d+))\.?[\t ]+/);
  if (article) {
    const value = article[3] ?? "";
    const ordinal = /^\d+$/.test(value)
      ? Number(value)
      : NUMBER_WORDS[value.toLowerCase()] ?? romanToNumber(value);
    if (Number.isFinite(ordinal) && ordinal > 0) {
      return createToken(article, article[1] ?? "", "article", ordinal, [ordinal], article[2] ?? "ARTICLE");
    }
  }

  const section = text.match(/^\s*((Section|section)\s+(\d+(?:\.\d+)*))\.?[\t ]+/);
  if (section) {
    const components = (section[3] ?? "").split(".").map(Number);
    return createToken(section, section[1] ?? "", "section", components.at(-1) ?? 0, components, section[2] ?? "Section");
  }

  const decimal = text.match(/^\s*(\d+(?:\.\d+)+)\.?[\t ]+/);
  if (decimal) {
    const components = (decimal[1] ?? "").split(".").map(Number);
    return createToken(decimal, decimal[1] ?? "", "decimal", components.at(-1) ?? 0, components, "");
  }

  const parenthetical = text.match(/^\s*\(([A-Za-z]+|\d+)\)[\t ]+/);
  if (parenthetical) {
    const value = parenthetical[1] ?? "";
    if (/^\d+$/.test(value)) {
      return createToken(parenthetical, `(${value})`, "decimal-item", Number(value), [Number(value)], "", ")");
    }
    const lower = value.toLowerCase();
    if (lower === "i") {
      return createToken(parenthetical, `(${value})`, "ambiguous-i", 1, [1], "", ")");
    }
    if (/^[ivxlcdm]+$/.test(lower) && lower.length > 1) {
      const kind = value === value.toUpperCase() ? "upper-roman" : "lower-roman";
      return createToken(parenthetical, `(${value})`, kind, romanToNumber(value), [romanToNumber(value)], "", ")");
    }
    const kind = value === value.toUpperCase() ? "upper-letter" : "lower-letter";
    return createToken(parenthetical, `(${value})`, kind, lettersToNumber(value), [lettersToNumber(value)], "", ")");
  }

  const delimiter = text.match(/^\s*(\d+|[A-Za-z]+)([.)])[\t ]+/);
  if (delimiter) {
    const value = delimiter[1] ?? "";
    if (/^\d+$/.test(value)) {
      return createToken(delimiter, `${value}${delimiter[2]}`, "decimal-item", Number(value), [Number(value)], "", delimiter[2] ?? "");
    }
    const lower = value.toLowerCase();
    const roman = /^[ivxlcdm]+$/.test(lower) && lower.length > 1;
    const kind: TokenKind = roman
      ? value === value.toUpperCase() ? "upper-roman" : "lower-roman"
      : value === value.toUpperCase() ? "upper-letter" : "lower-letter";
    const ordinal = roman ? romanToNumber(value) : lettersToNumber(value);
    return createToken(delimiter, `${value}${delimiter[2]}`, kind, ordinal, [ordinal], "", delimiter[2] ?? "");
  }

  return null;
}

export function resolveAmbiguousTokens(tokens: Array<ManualToken | null>, indents: Array<number | null>): void {
  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (token?.kind !== "ambiguous-i") continue;
    const indent = indents[index];
    const previous = tokens[index - 1];
    const next = tokens[index + 1];
    const samePreviousIndent = indents[index - 1] === indent;
    const sameNextIndent = indents[index + 1] === indent;
    if (samePreviousIndent && previous?.kind === "lower-letter" && previous.ordinal === 8) {
      token.kind = "lower-letter";
      token.ordinal = 9;
      token.components = [9];
    } else if (sameNextIndent && next?.kind === "lower-roman" && next.ordinal === 2) {
      token.kind = "lower-roman";
    } else if (samePreviousIndent && previous?.kind === "lower-roman") {
      token.kind = "lower-roman";
    }
  }
}
