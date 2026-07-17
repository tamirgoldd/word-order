import { readFile } from "node:fs/promises";
import { basename, resolve } from "node:path";
import { glob } from "node:fs/promises";
import { describe, expect, it } from "vitest";
import { createRepairPlan, repairDocument } from "../src/index.js";
import { compareTextModuloNumberTokens, validateDocxPackage } from "../src/testing.js";

const root = resolve(import.meta.dirname, "../../..");

describe("generated document corpus", async () => {
  const corpusFiles: string[] = [];
  for await (const file of glob("corpus/broken/*.docx", { cwd: root })) corpusFiles.push(resolve(root, file));

  it("contains at least twenty distinct broken-document fixtures", () => {
    expect(corpusFiles.length).toBeGreaterThanOrEqual(20);
  });

  for (const inputPath of corpusFiles.sort()) {
    const fixtureName = basename(inputPath, ".docx");
    it(`matches the golden plan and invariants for ${fixtureName}`, async () => {
      const bytes = new Uint8Array(await readFile(inputPath));
      const expectedPath = resolve(root, `corpus/golden/generated/${fixtureName}.plan.json`);
      const expected = JSON.parse(await readFile(expectedPath, "utf8")) as unknown;
      const plan = await createRepairPlan(bytes);
      expect({ status: plan.status, summary: plan.summary, anomalies: plan.anomalies, warnings: plan.warnings }).toEqual(expected);

      if (plan.status === "blocked") return;
      const result = await repairDocument(bytes, plan, { allowAnomalies: plan.anomalies.length > 0 });
      expect(await validateDocxPackage(result.bytes)).toEqual([]);
      expect((await compareTextModuloNumberTokens(bytes, result.bytes)).equal).toBe(true);
      const secondPlan = await createRepairPlan(result.bytes);
      expect(secondPlan.numbering).toHaveLength(0);
    });
  }
});
