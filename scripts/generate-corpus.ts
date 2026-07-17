import { mkdir, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { createRepairPlan } from "../packages/core/src/index.js";
import { createSyntheticDocx, type SyntheticParagraph } from "../packages/core/src/testing.js";

const root = resolve(import.meta.dirname, "..");
const brokenDirectory = resolve(root, "corpus/broken");
const goldenDirectory = resolve(root, "corpus/golden/generated");

const base: SyntheticParagraph[] = [
  { text: "ARTICLE I General" },
  { text: "Section 1.01 Purpose" },
  { text: "(a) First obligation" },
  { text: "(i) First detail" },
  { text: "(ii) Second detail" },
  { text: "(b) Second obligation" },
  { text: "Section 1.02 See Sections 1.01 and 1.03." },
  { text: "Section 1.03 Notices" }
];

type Mutation = (paragraphs: SyntheticParagraph[], variant: number) => SyntheticParagraph[];

const mutations: Array<[string, Mutation]> = [
  ["typed", (paragraphs) => paragraphs],
  ["sequence-jump", (paragraphs, variant) => paragraphs.map((paragraph, index) => index === 6
    ? { ...paragraph, text: `Section 1.${String(7 + variant).padStart(2, "0")} See Section 1.01.` }
    : paragraph)],
  ["missing-target", (paragraphs, variant) => paragraphs.map((paragraph, index) => index === 6
    ? { ...paragraph, text: `Section 1.02 See Section 8.${variant + 1}.` }
    : paragraph)],
  ["table-clause", (paragraphs, variant) => paragraphs.map((paragraph, index) => index === 2 + (variant % 2)
    ? { ...paragraph, table: true }
    : paragraph)],
  ["tracked-change", (paragraphs, variant) => paragraphs.map((paragraph, index) => index === variant % paragraphs.length
    ? { ...paragraph, tracked: "insert" as const }
    : paragraph)],
  ["paste-mismatch", (paragraphs, variant) => [
    ...paragraphs.slice(0, 6),
    { text: `Section ${4 + variant}.08 Pasted clause` },
    ...paragraphs.slice(6)
  ]]
];

await mkdir(brokenDirectory, { recursive: true });
await mkdir(goldenDirectory, { recursive: true });

for (const [mutationName, mutate] of mutations) {
  for (let variant = 0; variant < 4; variant += 1) {
    const name = `${mutationName}-${variant + 1}`;
    const bytes = await createSyntheticDocx(mutate(structuredClone(base), variant));
    const plan = await createRepairPlan(bytes);
    await writeFile(resolve(brokenDirectory, `${name}.docx`), bytes);
    await writeFile(resolve(goldenDirectory, `${name}.plan.json`), `${JSON.stringify({
      status: plan.status,
      summary: plan.summary,
      anomalies: plan.anomalies,
      warnings: plan.warnings
    }, null, 2)}\n`);
  }
}

console.log("Generated 24 synthetic corpus documents and plans.");
