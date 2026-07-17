import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { createRepairPlan, docxToFlatOpc, flatOpcToDocx, LegalDownError, repairDocument } from "../src/index.js";
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
    await expect(repairDocument(input, plan)).rejects.toMatchObject<Partial<LegalDownError>>({ code: "TRACKED_CHANGES" });
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
    await expect(repairDocument(input, plan)).rejects.toMatchObject<Partial<LegalDownError>>({ code: "CONFIRMATION_REQUIRED" });
    await expect(repairDocument(input, plan, { allowAnomalies: true })).resolves.toMatchObject({ changed: true });
  });

  it("rejects a plan for different source bytes", async () => {
    const first = await createSyntheticDocx([{ text: "Section 1.01 First" }]);
    const second = await createSyntheticDocx([{ text: "Section 2.01 Second" }]);
    const plan = await createRepairPlan(first);
    await expect(repairDocument(second, plan)).rejects.toMatchObject<Partial<LegalDownError>>({ code: "SOURCE_MISMATCH" });
  });

  it("adapts Word Flat OPC packages without changing visible content", async () => {
    const input = await createSyntheticDocx(basicParagraphs);
    const flatOpc = await docxToFlatOpc(input);
    const roundTrip = await flatOpcToDocx(flatOpc);
    expect(await validateDocxPackage(roundTrip)).toEqual([]);
    expect((await compareTextModuloNumberTokens(input, roundTrip)).equal).toBe(true);
  });
});
