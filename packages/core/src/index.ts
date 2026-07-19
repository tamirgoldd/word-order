import { WordOrderError } from "./errors.js";
import { analyzeFormatting, LEGAL_FORMATTING_PROFILE } from "./formatting.js";
import { inferNumbering, inferProfile } from "./infer.js";
import { inventoryDocument } from "./inventory.js";
import { fingerprint, loadDocx } from "./package.js";
import { buildRepairPlan } from "./plan.js";
import { rebuildDocx } from "./rebuild.js";
import type { RepairOptions, RepairPlan, RepairResult } from "./types.js";

export { WordOrderError } from "./errors.js";
export { docxToFlatOpc, flatOpcToDocx } from "./flat-opc.js";
export type * from "./types.js";

export async function createRepairPlan(bytes: Uint8Array): Promise<RepairPlan> {
  const docx = await loadDocx(bytes);
  const inventory = inventoryDocument(docx.documentXml, docx.numberingXml);
  const profile = inferProfile(inventory);
  const inference = inferNumbering(inventory, profile);
  const formatting = analyzeFormatting(docx.documentXml, inventory, inference.targets, LEGAL_FORMATTING_PROFILE);
  return buildRepairPlan({
    sourceFingerprint: fingerprint(bytes),
    inventory,
    profile,
    targets: inference.targets,
    numbering: inference.changes,
    anomalies: inference.anomalies,
    formattingProfile: LEGAL_FORMATTING_PROFILE,
    formatting: formatting.changes,
    formattingWarnings: formatting.warnings,
    trackedChanges: docx.hasTrackedChanges
  });
}

export async function repairDocument(
  bytes: Uint8Array,
  plan?: RepairPlan,
  options: RepairOptions = {}
): Promise<RepairResult> {
  const activePlan = plan ?? await createRepairPlan(bytes);
  if (activePlan.sourceFingerprint !== fingerprint(bytes)) {
    throw new WordOrderError("SOURCE_MISMATCH", "This repair plan belongs to a different version of the document.");
  }
  if (activePlan.status === "blocked") {
    throw new WordOrderError("TRACKED_CHANGES", activePlan.blockedReason ?? "This document cannot be repaired safely.");
  }
  if (activePlan.anomalies.length > 0 && !options.allowAnomalies) {
    throw new WordOrderError("CONFIRMATION_REQUIRED", "Review and confirm the flagged numbering anomalies before applying the repair.");
  }
  const docx = await loadDocx(bytes);
  const output = await rebuildDocx(docx, activePlan);
  const changed = output.length !== bytes.length || output.some((value, index) => value !== bytes[index]);
  return { bytes: output, plan: activePlan, changed };
}
