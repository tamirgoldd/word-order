import { useCallback, useRef, useState } from "react";
import {
  createRepairPlan,
  LegalDownError,
  repairDocument,
  type RepairPlan
} from "@legal-down/core";

type Stage = "idle" | "reading" | "review" | "repairing";

function fileSize(bytes: number): string {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function fixedName(name: string): string {
  return name.replace(/\.docx$/i, "") + ".fixed.docx";
}

function download(bytes: Uint8Array, name: string): void {
  const blob = new Blob([new Uint8Array(bytes)], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fixedName(name);
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1_000);
}

function PrivacyProof(): React.JSX.Element {
  return (
    <div className="privacy-proof" aria-label="Privacy guarantees">
      <span><i aria-hidden="true">01</i> Runs in this tab</span>
      <span><i aria-hidden="true">02</i> No upload</span>
      <span><i aria-hidden="true">03</i> Works offline</span>
    </div>
  );
}

function Summary({ plan }: { plan: RepairPlan }): React.JSX.Element {
  const stats = [
    [plan.summary.numberingConversions, "numbers repaired"],
    [plan.formatting.length, "formatting repairs"],
    [plan.summary.crossReferencesConverted, "references linked"],
    [plan.summary.brokenReferences, "broken references"],
    [plan.summary.anomalies + plan.warnings.length, "items to review"]
  ] as const;
  return (
    <div className="summary-grid">
      {stats.map(([value, label]) => (
        <div className={label.includes("broken") && value ? "stat warning-stat" : "stat"} key={label}>
          <strong>{value}</strong>
          <span>{label}</span>
        </div>
      ))}
    </div>
  );
}

function PlanReview({ plan }: { plan: RepairPlan }): React.JSX.Element {
  const formattingGroups = [...plan.formatting.reduce((groups, change) => {
    const list = groups.get(change.category) ?? [];
    list.push(change);
    groups.set(change.category, list);
    return groups;
  }, new Map<string, typeof plan.formatting>()).entries()];
  return (
    <div className="plan-sections">
      {plan.numbering.length > 0 && (
        <section className="review-card">
          <div className="section-heading">
            <div><span className="eyebrow">Numbering plan</span><h3>Typed labels become native Word lists</h3></div>
            <span className="count">{plan.numbering.length}</span>
          </div>
          <div className="change-list">
            {plan.numbering.slice(0, 12).map((change) => (
              <div className="change-row" key={change.paragraphIndex}>
                <code className="old-number">{change.oldNumber}</code>
                <span className="arrow" aria-hidden="true">→</span>
                <code className="new-number">{change.newNumber}</code>
                <span className="preview">{change.textPreview || "Untitled paragraph"}</span>
                <span className={change.confidence === "high" ? "confidence" : "confidence review"}>{change.confidence}</span>
              </div>
            ))}
          </div>
          {plan.numbering.length > 12 && <p className="more">+ {plan.numbering.length - 12} more changes in the downloaded report</p>}
        </section>
      )}

      {plan.formatting.length > 0 && (
        <section className="review-card">
          <div className="section-heading">
            <div><span className="eyebrow">Formatting plan</span><h3>Inconsistent styling becomes a restrained legal-document system</h3></div>
            <span className="count">{plan.formatting.length}</span>
          </div>
          <div className="formatting-list">
            {formattingGroups.map(([category, changes]) => {
              const example = changes[0];
              return (
                <div className="formatting-row" key={category}>
                  <strong>{category.replace("-", " ")}</strong>
                  <span>{example?.oldValue} <i aria-hidden="true">→</i> {example?.newValue}</span>
                  <small>{changes.length} {changes.length === 1 ? "repair" : "repairs"}</small>
                </div>
              );
            })}
          </div>
        </section>
      )}

      {plan.crossReferences.length > 0 && (
        <section className="review-card">
          <div className="section-heading">
            <div><span className="eyebrow">Cross-reference audit</span><h3>Live links and dead ends</h3></div>
            <span className="count">{plan.crossReferences.length}</span>
          </div>
          <div className="reference-list">
            {plan.crossReferences.map((reference, index) => (
              <div className="reference-row" key={`${reference.paragraphIndex}-${reference.start}-${index}`}>
                <span className={reference.status === "resolved" ? "status-dot resolved" : "status-dot unresolved"} aria-hidden="true" />
                <span><code>{reference.display}</code> {reference.status === "resolved" ? `links to ${reference.targetNumber}` : reference.reason}</span>
                <small>¶ {reference.paragraphIndex + 1}</small>
              </div>
            ))}
          </div>
        </section>
      )}

      {plan.anomalies.length > 0 && (
        <section className="review-card anomaly-card">
          <div className="section-heading">
            <div><span className="eyebrow">Your decision</span><h3>We found numbering that is not safe to guess</h3></div>
          </div>
          {plan.anomalies.map((anomaly) => (
            <div className="anomaly" key={anomaly.paragraphIndex}>
              <span>¶ {anomaly.paragraphIndex + 1}</span>
              <p>{anomaly.message}</p>
              {anomaly.proposedNumber && <code>{anomaly.oldNumber} → {anomaly.proposedNumber}</code>}
            </div>
          ))}
        </section>
      )}


      {plan.warnings.length > 0 && (
        <section className="review-card warning-card">
          <div className="section-heading">
            <div><span className="eyebrow">Editorial review</span><h3>Wording we deliberately did not rewrite</h3></div>
            <span className="count">{plan.warnings.length}</span>
          </div>
          {plan.warnings.map((warning, index) => (
            <div className="anomaly" key={`${warning.paragraphIndex}-${warning.code}-${index}`}>
              <span>¶ {warning.paragraphIndex + 1}</span>
              <p>{warning.message}</p>
            </div>
          ))}
        </section>
      )}
    </div>
  );
}

export function App(): React.JSX.Element {
  const inputRef = useRef<HTMLInputElement>(null);
  const [stage, setStage] = useState<Stage>("idle");
  const [file, setFile] = useState<File | null>(null);
  const [bytes, setBytes] = useState<Uint8Array | null>(null);
  const [plan, setPlan] = useState<RepairPlan | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [confirmed, setConfirmed] = useState(false);
  const [dragging, setDragging] = useState(false);

  const scan = useCallback(async (nextFile: File) => {
    setError(null);
    setConfirmed(false);
    if (!nextFile.name.toLowerCase().endsWith(".docx")) {
      setError("Choose a modern Word .docx file. Legacy .doc files are not supported.");
      return;
    }
    if (nextFile.size > 200 * 1024 * 1024) {
      setError("This document is over 200 MB and is too large to process safely in a browser tab.");
      return;
    }
    setStage("reading");
    try {
      const nextBytes = new Uint8Array(await nextFile.arrayBuffer());
      const nextPlan = await createRepairPlan(nextBytes);
      setFile(nextFile);
      setBytes(nextBytes);
      setPlan(nextPlan);
      setStage("review");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "The document could not be read.");
      setStage("idle");
    }
  }, []);

  const apply = async () => {
    if (!file || !bytes || !plan) return;
    setError(null);
    setStage("repairing");
    try {
      const result = await repairDocument(bytes, plan, { allowAnomalies: confirmed });
      if (result.changed) download(result.bytes, file.name);
      else setError("This document is already structurally clean. No new file was needed.");
    } catch (cause) {
      setError(cause instanceof LegalDownError ? cause.message : "The repair could not be completed safely.");
    } finally {
      setStage("review");
    }
  };

  const reset = () => {
    setStage("idle");
    setFile(null);
    setBytes(null);
    setPlan(null);
    setError(null);
    setConfirmed(false);
    if (inputRef.current) inputRef.current.value = "";
  };

  const busy = stage === "reading" || stage === "repairing";
  const blocked = plan?.status === "blocked";
  const confirmationNeeded = Boolean(plan?.anomalies.length);

  return (
    <div className="page-shell">
      <header className="site-header">
        <a className="brand" href="." aria-label="Legal Down home"><span className="brand-mark">LD</span><span>Legal Down</span></a>
        <nav>
          <a href="#how-it-works">How it works</a>
          <a href="https://github.com/tamirgoldd/legal-down" target="_blank" rel="noreferrer">Source ↗</a>
        </nav>
      </header>

      <main>
        <section className={plan ? "hero compact" : "hero"}>
          <div className="hero-copy">
            <p className="kicker"><span /> Free · open source · local-first</p>
            <h1>Your Word draft broke.<br /><em>Fix the structure.</em></h1>
            <p className="lede">Repair mangled legal numbering, dead cross-references, and chaotic formatting in one pass. Get back a professional Word document that keeps working when the draft changes.</p>
            <PrivacyProof />
          </div>
          {!plan && <div className="document-motif" aria-hidden="true"><div className="paper back" /><div className="paper"><b>ARTICLE IV</b><span>Section 4.01</span><span className="broken-line">8.3</span><span>(a)</span><i>REF</i></div></div>}
        </section>

        <section className="workspace" aria-live="polite">
          {!plan ? (
            <div
              className={dragging ? "dropzone dragging" : "dropzone"}
              onDragEnter={(event) => { event.preventDefault(); setDragging(true); }}
              onDragOver={(event) => event.preventDefault()}
              onDragLeave={(event) => { if (event.currentTarget === event.target) setDragging(false); }}
              onDrop={(event) => {
                event.preventDefault();
                setDragging(false);
                const dropped = event.dataTransfer.files[0];
                if (dropped) void scan(dropped);
              }}
            >
              <input ref={inputRef} type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" onChange={(event) => {
                const selected = event.target.files?.[0];
                if (selected) void scan(selected);
              }} />
              <div className="upload-glyph" aria-hidden="true">DOCX<span>↑</span></div>
              <h2>{stage === "reading" ? "Reading the document locally…" : "Drop the broken agreement here"}</h2>
              <p>{stage === "reading" ? "Inventorying paragraphs, list definitions, and references." : "or choose a .docx file — nothing is uploaded"}</p>
              <button className="primary" type="button" disabled={busy} onClick={() => inputRef.current?.click()}>{busy ? "Scanning…" : "Choose Word document"}</button>
              <small>Your original file is never overwritten.</small>
            </div>
          ) : (
            <div className="review-workspace">
              <div className="file-bar">
                <div className="file-icon">W</div>
                <div><strong>{file?.name}</strong><span>{file ? fileSize(file.size) : ""} · scanned on this device</span></div>
                <button className="text-button" type="button" onClick={reset}>Choose another</button>
              </div>

              {blocked && <div className="blocking-message"><strong>Repair paused</strong><p>{plan.blockedReason}</p></div>}
              {!blocked && <Summary plan={plan} />}
              <PlanReview plan={plan} />

              {!blocked && plan.status !== "clean" && (
                <div className="apply-bar">
                  <div>
                    {confirmationNeeded ? (
                      <label className="confirmation"><input type="checkbox" checked={confirmed} onChange={(event) => setConfirmed(event.target.checked)} /><span>I reviewed the flagged changes and approve the proposed numbering.</span></label>
                    ) : <p><strong>Ready to repair.</strong><br />Your original stays untouched.</p>}
                  </div>
                  <button className="primary download-button" type="button" disabled={busy || (confirmationNeeded && !confirmed)} onClick={() => void apply()}>{stage === "repairing" ? "Building document…" : "Download fixed .docx"}</button>
                </div>
              )}
              {plan.status === "clean" && <div className="clean-message"><strong>No automatic repairs remain.</strong><span>{plan.warnings.length ? `${plan.warnings.length} editorial ${plan.warnings.length === 1 ? "warning remains" : "warnings remain"} for human review.` : "The document already uses consistent native numbering, formatting, and references."}</span></div>}
            </div>
          )}
          {error && <div className="error-message" role="alert">{error}</div>}
        </section>

        {!plan && (
          <>
            <section className="how" id="how-it-works">
              <div><span className="eyebrow">What changes</span><h2>Not a cosmetic patch.<br />A structural repair.</h2></div>
              <div className="principles">
                <article><span>1</span><h3>Inventory</h3><p>Reads typed numbers, Word lists, indentation, styles, tables, and references.</p></article>
                <article><span>2</span><h3>Review</h3><p>Shows every proposed change and refuses ambiguous structure until you confirm.</p></article>
                <article><span>3</span><h3>Rebuild</h3><p>Writes native multilevel lists, bookmarks, and live Word REF fields.</p></article>
              </div>
            </section>
            <section className="faq">
              <span className="eyebrow">Plain answers</span>
              <details><summary>Does my contract leave the computer?</summary><p>No. The page has no document API or upload endpoint. Parsing and rebuilding happen in this browser tab, and the app can run after you disconnect from the internet.</p></details>
              <details><summary>Will the numbers still work after someone edits the file?</summary><p>Yes. Legal Down creates Word's native multilevel numbering rather than hardcoded text, then uses bookmarks and REF fields for cross-references.</p></details>
              <details><summary>Does it rewrite contract language?</summary><p>No. It normalizes styles, spacing, margins, highlighting, and structure. Suspiciously long run-on paragraphs are flagged for human review without changing their wording.</p></details>
              <details><summary>What happens to tracked changes?</summary><p>Legal Down stops and asks you to accept or reject them first. It does not attempt a risky repair through revision markup.</p></details>
            </section>
          </>
        )}
      </main>

      <footer><span>Legal Down is open source software, not legal advice.</span><a href="https://github.com/tamirgoldd/legal-down">MIT licensed · GitHub ↗</a></footer>
    </div>
  );
}
