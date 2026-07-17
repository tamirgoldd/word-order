#!/usr/bin/env node
import { access, readFile, writeFile } from "node:fs/promises";
import { basename, resolve } from "node:path";
import { createRepairPlan, LegalDownError, repairDocument, type RepairPlan } from "@legal-down/core";

interface Arguments {
  command: "fix" | "scan";
  input: string;
  output: string | null;
  report: string | null;
  allowAnomalies: boolean;
  force: boolean;
}

const USAGE = `Legal Down — repair legal Word structure and formatting locally

Usage:
  legal-down scan <input.docx> [--report plan.json]
  legal-down fix <input.docx> -o <output.docx> [--report plan.json]
                 [--allow-anomalies] [--force]

The file is processed on this machine. No network request is made.`;

function valueAfter(args: string[], short: string, long: string): string | null {
  const index = args.findIndex((argument) => argument === short || argument === long);
  return index >= 0 ? args[index + 1] ?? null : null;
}

function parseArguments(args: string[]): Arguments {
  if (args.includes("--help") || args.includes("-h")) {
    console.log(USAGE);
    process.exit(0);
  }
  const command = args[0];
  const input = args[1];
  if ((command !== "scan" && command !== "fix") || !input) throw new Error(USAGE);
  const output = valueAfter(args, "-o", "--output");
  if (command === "fix" && !output) throw new Error(`fix requires -o <output.docx>.\n\n${USAGE}`);
  return {
    command,
    input: resolve(input),
    output: output ? resolve(output) : null,
    report: valueAfter(args, "--report", "--report") ? resolve(valueAfter(args, "--report", "--report") ?? "") : null,
    allowAnomalies: args.includes("--allow-anomalies"),
    force: args.includes("--force")
  };
}

function humanPlan(plan: RepairPlan): string {
  const lines = [
    `${plan.status.toUpperCase()} — ${plan.summary.paragraphs} paragraphs scanned`,
    `${plan.summary.numberingConversions} numbering conversions; ${plan.formatting.length} formatting repairs; ${plan.summary.crossReferencesConverted} live references; ${plan.summary.brokenReferences} broken references; ${plan.summary.anomalies} anomalies`
  ];
  if (plan.blockedReason) lines.push(`\nBlocked: ${plan.blockedReason}`);
  if (plan.anomalies.length) {
    lines.push("\nReview required:");
    for (const anomaly of plan.anomalies) lines.push(`  ¶${anomaly.paragraphIndex + 1}: ${anomaly.message}`);
  }
  if (plan.warnings.length) {
    lines.push("\nWarnings:");
    for (const warning of plan.warnings) lines.push(`  ¶${warning.paragraphIndex + 1}: ${warning.message}`);
  }
  return lines.join("\n");
}

async function exists(path: string): Promise<boolean> {
  try {
    await access(path);
    return true;
  } catch {
    return false;
  }
}

async function main(): Promise<void> {
  const args = parseArguments(process.argv.slice(2));
  if (!args.input.toLowerCase().endsWith(".docx")) throw new Error("Legal Down supports .docx files only.");
  const input = new Uint8Array(await readFile(args.input));
  const plan = await createRepairPlan(input);
  console.log(humanPlan(plan));
  if (args.report) await writeFile(args.report, `${JSON.stringify(plan, null, 2)}\n`, { flag: args.force ? "w" : "wx" });
  if (args.command === "scan") return;
  if (!args.output) return;
  if (resolve(args.output) === resolve(args.input)) throw new Error("Refusing to overwrite the input document. Choose a different output path.");
  if (!args.force && await exists(args.output)) throw new Error(`Refusing to overwrite ${basename(args.output)}. Pass --force to replace it.`);
  const result = await repairDocument(input, plan, { allowAnomalies: args.allowAnomalies });
  if (!result.changed) {
    console.log("\nNothing needed repair; no output file was written.");
    return;
  }
  await writeFile(args.output, result.bytes, { flag: args.force ? "w" : "wx" });
  console.log(`\nWrote ${args.output}`);
}

main().catch((error: unknown) => {
  if (error instanceof LegalDownError) console.error(`${error.code}: ${error.message}`);
  else console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
