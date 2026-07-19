import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { createRepairPlan, docxToFlatOpc, flatOpcToDocx, WordOrderError, repairDocument } from "../src/index.js";
import {
  compareTextModuloNumberTokens,
  createSyntheticDocx,
  validateDocxPackage
} from "../src/testing.js";

const basicParagraphs = [
  { text: "ARTICLE I General" },
  { text: "Section 1.01 Purpose" },
  { text: "(a) First obligation" },
  { text: "(i) First detail" },
  { text: "(ii) Second detail" },
  { text: "(b) Second obligation" },
  { text: "Section 1.02 Notices are governed by Section 1.01 and Section 9.99." }
];

describe("repair planning", () => {
  it("inventories manual numbering, resolves contextual (i), and audits references", async () => {
    const input = await createSyntheticDocx(basicParagraphs);
    const plan = await createRepairPlan(input);

    expect(plan.status).toBe("ready");
    expect(plan.summary).toEqual({
      paragraphs: 7,
      numberedParagraphs: 7,
      numberingConversions: 7,
      crossReferencesConverted: 1,
      brokenReferences: 1,
      anomalies: 0
    });
    expect(plan.inventory[3]?.manual?.kind).toBe("lower-roman");
    expect(plan.crossReferences.find((reference) => reference.display === "1.01")?.status).toBe("resolved");
    expect(plan.crossReferences.find((reference) => reference.display === "9.99")?.status).toBe("unresolved");
  });

  it("refuses a context-free (i) rather than guessing", async () => {
    const input = await createSyntheticDocx([{ text: "(i) Ambiguous standalone paragraph" }]);
    const plan = await createRepairPlan(input);

    expect(plan.status).toBe("needs-confirmation");
    expect(plan.anomalies[0]?.code).toBe("ambiguous-token");
    expect(plan.numbering).toHaveLength(0);
  });

  it("flags sequence jumps with a proposed repair", async () => {
    const input = await createSyntheticDocx([
      { text: "ARTICLE I General" },
      { text: "Section 1.01 First" },
      { text: "Section 1.08 Pasted from elsewhere" },
      { text: "Section 1.09 Next" }
    ]);
    const plan = await createRepairPlan(input);

    expect(plan.status).toBe("needs-confirmation");
    expect(plan.anomalies.map((anomaly) => anomaly.proposedNumber)).toContain("Section 1.02");
  });

  it("blocks tracked changes anywhere in the Word package", async () => {
    const input = await createSyntheticDocx([
      { text: "ARTICLE I Inserted", tracked: "insert" },
      { text: "Section 1.01 Body" }
    ]);
    const plan = await createRepairPlan(input);

    expect(plan.status).toBe("blocked");
    await expect(repairDocument(input, plan)).rejects.toMatchObject<Partial<WordOrderError>>({ code: "TRACKED_CHANGES" });
  });

  it("includes numbered paragraphs inside tables", async () => {
    const input = await createSyntheticDocx([
      { text: "ARTICLE I General" },
      { text: "Section 1.01 In a cell", table: true }
    ]);
    const plan = await createRepairPlan(input);
    expect(plan.inventory[1]?.inTable).toBe(true);
    expect(plan.summary.numberedParagraphs).toBe(2);
  });
});

describe("OOXML rebuild", () => {
  it("normalizes chaotic legal formatting and repairs flat section numbering without rewriting prose", async () => {
    const input = await createSyntheticDocx([
      { text: "SERVICES agreement" },
      { text: "(the \"Agreement\")" },
      { text: "This agreement is between ACME HOLDINGS, LLC and the Provider." },
      { text: "1.  DEFINITIONS" },
      { text: "A short body paragraph governed by Section 1." },
      { text: "3.  payment terms" },
      { text: "Company shall pay $[AMOUNT] within thirty days." },
      { text: "3.  Term and Termination" },
      { text: "Either party may terminate on written notice." },
      { text: "2.  CONFIDENTIALITY" },
      { text: `The Provider acknowledges that confidential information must remain protected and agrees not to disclose it to any third party and this obligation continues after termination and applies to all business records and materials and the Provider must use reasonable safeguards and notify the Company promptly of any suspected breach or unauthorized disclosure and must cooperate fully with every investigation and remediation effort requested by the Company.` },
      { text: "Company: _______________________" }
    ]);
    const zip = await JSZip.loadAsync(input);
    let document = await zip.file("word/document.xml")?.async("string") ?? "";
    document = document
      .replace("<w:p><w:r><w:t xml:space=\"preserve\">SERVICES agreement</w:t></w:r></w:p>", "<w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Comic Sans MS\"/><w:b/><w:sz w:val=\"56\"/></w:rPr><w:t xml:space=\"preserve\">SERVICES agreement</w:t></w:r></w:p>")
      .replace("<w:p><w:r><w:t xml:space=\"preserve\">This agreement is between ACME HOLDINGS, LLC and the Provider.</w:t></w:r></w:p>", "<w:p><w:r><w:rPr><w:rFonts w:ascii=\"Arial\"/><w:sz w:val=\"18\"/><w:i/></w:rPr><w:t xml:space=\"preserve\">This agreement is between ACME HOLDINGS, LLC and the Provider.</w:t></w:r></w:p>")
      .replace("<w:p><w:r><w:t xml:space=\"preserve\">A short body paragraph governed by Section 1.</w:t></w:r></w:p>", "<w:p><w:pPr><w:ind w:left=\"1400\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Georgia\"/><w:sz w:val=\"28\"/><w:u w:val=\"single\"/></w:rPr><w:t xml:space=\"preserve\">A short body paragraph governed by Section 1.</w:t></w:r></w:p>")
      .replace("<w:p><w:r><w:t xml:space=\"preserve\">Company shall pay $[AMOUNT] within thirty days.</w:t></w:r></w:p>", "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Courier New\"/><w:sz w:val=\"30\"/><w:highlight w:val=\"yellow\"/></w:rPr><w:t xml:space=\"preserve\">Company shall pay $[AMOUNT] within thirty days.</w:t></w:r></w:p>")
      .replace("<w:sectPr/>", "<w:sectPr><w:pgMar w:top=\"400\" w:right=\"2200\" w:bottom=\"1600\" w:left=\"500\"/></w:sectPr>");
    zip.file("word/document.xml", document);
    const broken = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });

    const plan = await createRepairPlan(broken);
    expect(plan.profile.topLevelKind).toBe("flat-section");
    expect(plan.numbering.map((change) => change.newNumber)).toEqual(["1.", "2.", "3.", "4."]);
    expect(plan.crossReferences.find((reference) => reference.display === "1")?.status).toBe("resolved");
    expect([...new Set(plan.formatting.map((change) => change.category))]).toEqual(expect.arrayContaining(["font-family", "font-size", "alignment", "indent", "highlight", "margins"]));
    expect(plan.warnings.some((warning) => warning.code === "long-paragraph")).toBe(true);

    const repaired = await repairDocument(broken, plan, { allowAnomalies: true });
    const output = await JSZip.loadAsync(repaired.bytes);
    const outputDocument = await output.file("word/document.xml")?.async("string") ?? "";
    const outputStyles = await output.file("word/styles.xml")?.async("string") ?? "";
    const outputNumbering = await output.file("word/numbering.xml")?.async("string") ?? "";
    expect(outputDocument).not.toMatch(/Comic Sans|Arial|Georgia|Courier New|<w:highlight\b/);
    expect(outputDocument).toContain('w:pStyle w:val="LDTitle"');
    expect(outputDocument).toContain('w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"');
    expect(outputStyles).toContain('w:styleId="LDBody"');
    expect(outputNumbering).toContain('w:lvlText w:val="%1."');
    expect((await compareTextModuloNumberTokens(broken, repaired.bytes)).equal).toBe(true);

    const secondPlan = await createRepairPlan(repaired.bytes);
    const second = await repairDocument(repaired.bytes, secondPlan);
    expect(secondPlan.status).toBe("clean");
    expect(secondPlan.numbering).toHaveLength(0);
    expect(secondPlan.formatting).toHaveLength(0);
    expect(second.changed).toBe(false);
  });

  it("emits native numbering, REF fields, bookmarks, styles, and updateFields", async () => {
    const input = await createSyntheticDocx(basicParagraphs);
    const plan = await createRepairPlan(input);
    const result = await repairDocument(input, plan);
    const zip = await JSZip.loadAsync(result.bytes);
    const document = await zip.file("word/document.xml")?.async("string");
    const numbering = await zip.file("word/numbering.xml")?.async("string");
    const styles = await zip.file("word/styles.xml")?.async("string");
    const settings = await zip.file("word/settings.xml")?.async("string");

    expect(result.changed).toBe(true);
    expect(document).toContain("<w:numPr>");
    expect(document).toContain("_LDRef_2");
    expect(document).toContain(" REF _LDRef_2 \\r \\h ");
    expect(numbering).toContain('w:lvlText w:val="ARTICLE %1"');
    expect(numbering).toContain('w:lvlText w:val="Section %1.%2"');
    expect(numbering).toContain("<w:isLgl/>");
    expect(styles).toContain('w:styleId="LDSection"');
    expect(settings).toContain('w:updateFields w:val="true"');
    expect(await validateDocxPackage(result.bytes)).toEqual([]);
    expect((await compareTextModuloNumberTokens(input, result.bytes)).equal).toBe(true);
  });

  it("converts a reference split across differently formatted Word runs", async () => {
    const input = await createSyntheticDocx([
      { text: "Section 1.01 Target" },
      { text: "Subject to Section 1.01." }
    ]);
    const zip = await JSZip.loadAsync(input);
    const document = await zip.file("word/document.xml")?.async("string");
    zip.file("word/document.xml", document?.replace(
      "<w:r><w:t xml:space=\"preserve\">Subject to Section 1.01.</w:t></w:r>",
      "<w:r><w:t>Subject to Section 1.</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>01</w:t></w:r><w:r><w:t>.</w:t></w:r>"
    ) ?? "");
    const splitInput = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    const result = await repairDocument(splitInput);
    const output = await JSZip.loadAsync(result.bytes);
    const outputDocument = await output.file("word/document.xml")?.async("string");

    expect(outputDocument).toContain(" REF _LDRef_1 \\r \\h ");
    expect((await compareTextModuloNumberTokens(splitInput, result.bytes)).equal).toBe(true);
  });

  it("is idempotent after repair", async () => {
    const input = await createSyntheticDocx(basicParagraphs);
    const first = await repairDocument(input);
    const secondPlan = await createRepairPlan(first.bytes);
    const second = await repairDocument(first.bytes, secondPlan);

    expect(secondPlan.status).toBe("clean");
    expect(secondPlan.numbering).toHaveLength(0);
    expect(secondPlan.crossReferences).toHaveLength(0);
    expect(second.changed).toBe(false);
    expect(second.bytes).toEqual(first.bytes);
  });

  it("keeps renumbering after a native numbered paragraph is inserted", async () => {
    const input = await createSyntheticDocx(basicParagraphs);
    const repaired = await repairDocument(input);
    const zip = await JSZip.loadAsync(repaired.bytes);
    const document = await zip.file("word/document.xml")?.async("string");
    expect(document).toBeTruthy();
    const sectionParagraph = document
      ? [...document.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].find((match) => match[0].includes("Notices are governed"))?.[0]
      : undefined;
    const numId = sectionParagraph?.match(/<w:numId w:val="(\d+)"\/>/)?.[1];
    expect(numId).toBeTruthy();
    const inserted = `<w:p><w:pPr><w:pStyle w:val="LDSection"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>Inserted section</w:t></w:r></w:p>`;
    zip.file("word/document.xml", document?.replace(sectionParagraph ?? "", `${inserted}${sectionParagraph}`) ?? "");
    const edited = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    const editedPlan = await createRepairPlan(edited);

    expect(editedPlan.status).toBe("clean");
    expect(editedPlan.targets.find((target) => target.paragraphIndex === 6)?.number).toBe("Section 1.02");
    expect(editedPlan.targets.find((target) => target.paragraphIndex === 7)?.number).toBe("Section 1.03");
  });

  it("requires explicit confirmation for anomalous plans", async () => {
    const input = await createSyntheticDocx([
      { text: "Section 1.01 First" },
      { text: "Section 1.09 Wrong" }
    ]);
    const plan = await createRepairPlan(input);
    await expect(repairDocument(input, plan)).rejects.toMatchObject<Partial<WordOrderError>>({ code: "CONFIRMATION_REQUIRED" });
    await expect(repairDocument(input, plan, { allowAnomalies: true })).resolves.toMatchObject({ changed: true });
  });

  it("rejects a plan for different source bytes", async () => {
    const first = await createSyntheticDocx([{ text: "Section 1.01 First" }]);
    const second = await createSyntheticDocx([{ text: "Section 2.01 Second" }]);
    const plan = await createRepairPlan(first);
    await expect(repairDocument(second, plan)).rejects.toMatchObject<Partial<WordOrderError>>({ code: "SOURCE_MISMATCH" });
  });

  it("adapts Word Flat OPC packages without changing visible content", async () => {
    const input = await createSyntheticDocx(basicParagraphs);
    const flatOpc = await docxToFlatOpc(input);
    expect(flatOpc).not.toContain('pkg:name="/[Content_Types].xml"');
    const roundTrip = await flatOpcToDocx(flatOpc);
    expect(await validateDocxPackage(roundTrip)).toEqual([]);
    expect((await compareTextModuloNumberTokens(input, roundTrip)).equal).toBe(true);
  });
});
