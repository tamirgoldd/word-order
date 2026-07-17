# Legal-Down — Build Specification

## What this is (short version)

A free, open-source, 100% client-side tool that **repairs broken numbering and cross-references in legal Word documents**. A lawyer opens a mangled 50-page agreement, clicks "Fix," and gets back the same document with clean native Word multilevel numbering (ARTICLE I → Section 1.01 → (a) → (i)) and all textual cross-references ("subject to Section 4.02") converted into live Word REF fields that update automatically. No cloud, no upload, no account — the document never leaves the machine. Two surfaces, one engine: a drag-and-drop web app (discovery/front door) and a Word add-in (the tool people keep).

**This is NOT a markdown editor.** The original markdown-drafting concept is deferred to a possible v2. v1 is repair-only, because it requires zero behavior change from lawyers and works mid-negotiation on any document, including opposing counsel's redlines.

---

## Core principles (hard constraints — do not violate)

1. **100% local.** All parsing and rebuilding happens in the browser (or in Word's JS runtime). No document byte ever leaves the client. No telemetry on document content. No backend.
2. **Native Word output, never static strings.** The engine must emit real OOXML multilevel list definitions (`numbering.xml` / `abstractNum`) and real `REF` fields anchored to bookmarks — NOT hardcoded number strings. A document fixed by this tool must keep renumbering correctly when someone later edits it in plain Word. Static strings would make the output *more* fragile than what the user started with.
3. **Zero content loss.** The plain-text content of the document (excluding the numbering tokens themselves) must be byte-identical before and after repair. This is a tested invariant, not a goal.
4. **Graceful failure.** If the engine cannot confidently infer structure, it must say so and do nothing, rather than guess and corrupt. A tool that silently mangles one contract loses the entire audience.
5. **Boring and dependency-light.** Minimal dependencies, no framework churn. This tool must be maintainable for 5+ years.

---

## Architecture

Monorepo (pnpm workspaces), TypeScript throughout:

```
legal-down/
  packages/
    core/        # THE PRODUCT. Pure TS library, zero DOM deps. docx in → docx out.
    cli/         # thin node wrapper around core, for testing & power users
    web/         # drag-and-drop web app (static site, Vite). Uses core.
    addin/       # Office.js Word add-in taskpane. Uses core.
  corpus/        # test documents + golden files (see Testing)
```

**Build order: core → cli → web → addin.** Do not touch UI until the engine passes the corpus tests. The engine is 80% of the difficulty and 100% of the value.

### Stack

- **TypeScript**, strict mode
- **JSZip** for .docx (it's a zip of XML)
- **fast-xml-parser** or direct XML manipulation — but preserve unknown XML untouched (round-trip fidelity matters; do not re-serialize parts you didn't change)
- **Vitest** for tests
- **Vite** + plain React (or Preact) for web app — keep it a single static page
- **Office.js** for the add-in
- No docx.js needed — we manipulate OOXML directly, which gives full control over numbering definitions and fields that high-level libraries don't expose properly

---

## Package: `core` — the repair engine

### Pipeline

```
.docx bytes
  → unzip, parse document.xml, numbering.xml, styles.xml
  → PHASE 1: INVENTORY   — classify every paragraph
  → PHASE 2: INFERENCE   — build the intended hierarchy tree
  → PHASE 3: PLAN        — produce a repair plan (diffable, user-reviewable)
  → PHASE 4: REBUILD     — emit new OOXML
  → rezip → .docx bytes + repair report
```

### Phase 1 — Inventory

For each body paragraph, detect how it's numbered. There are three populations, and real documents mix all three (that's exactly why they're broken):

1. **Native auto-numbering**: paragraph has `w:numPr` (numId + ilvl) referencing `numbering.xml`. Resolve the actual rendered number by simulating Word's counter logic (including `w:lvlOverride`, `w:startOverride`, restart behavior, and the `w:isLgl` flag).
2. **Manually typed numbers**: paragraph text *starts with* a number token. Detect via ordered pattern set:
   - `ARTICLE I` / `ARTICLE 1` / `Article One` (word + roman/decimal/spelled)
   - `Section 1.01` / `Section 1.1` / `1.01` / `1.1` / `1.1.1` (decimal chains, with/without label word, with/without trailing period)
   - `(a)` `(b)` … lowercase letters; `(A)` uppercase
   - `(i)` `(ii)` … lowercase roman; `(I)` uppercase roman
   - `1)` `a.` `i.` and other delimiter variants
   - Tab or space separator between token and text
3. **Unnumbered**: headings by style, recitals, definitions, plain body text — leave alone.

**Critical ambiguity — `(i)`**: could be roman-numeral level or letter level (comes after `(h)`). Resolve by context: if the previous sibling at the same indent/level is `(h)`, it's a letter; if the sequence is `(i)(ii)(iii)`, it's roman. Never resolve `(i)` in isolation.

Also record per paragraph: indentation (both `w:ind` values and style-inherited), style name, whether it sits inside a table, and whether it carries tracked-changes markup.

### Phase 2 — Inference

Build a tree of the *intended* hierarchy:

- Cluster the detected schemes into levels (e.g., ARTICLE-roman = level 0, Section-x.yy = level 1, (a) = level 2, (i) = level 3).
- Use **sequence continuity** as the primary signal, indentation as secondary, style as tertiary. A `(c)` following a `(b)` at similar indent is the same level even if Word's list data disagrees — Word's own data is what's broken.
- Detect **restarts**: `(a)` sequences restart under each new Section. Model numbering as restart-per-parent for sub-levels (standard legal convention), but verify against the document's own dominant pattern rather than assuming.
- Detect the document's **scheme profile**: does it use `Section 1.01` (two-digit, big-law style) or `1.1`? ARTICLE or Article? Preserve the document's existing conventions — never impose a house style the document didn't have.
- Flag anomalies instead of guessing: a paragraph numbered `8.3` appearing between `2.4` and `2.5` (classic paste-from-another-matter) gets flagged in the plan as "renumber to 2.5?" — the user confirms.

Output: a tree where every numbered paragraph has `(level, ordinal, schemeProfile)` plus a list of anomalies.

### Phase 3 — Plan

Produce a machine-readable + human-readable repair plan before touching anything:

- every paragraph: old number token → new computed number
- every detected cross-reference: text → target paragraph → resolved new number
- every anomaly requiring user confirmation
- summary stats (paragraphs renumbered, xrefs converted, anomalies)

The UIs render this as a review screen. **Nothing is applied without the plan.** This is also what makes the engine testable: plans are golden-filed.

### Phase 4 — Rebuild

1. **Create one clean `abstractNum`** implementing the inferred scheme profile. Legal conventions in OOXML:
   - Level 0: `numFmt=upperRoman`, `lvlText="ARTICLE %1"`
   - Level 1: `numFmt=decimalZero` (produces 01, 02…), `lvlText="Section %1.%2"`, and set `w:isLgl` where the profile calls for full-context legal numbering
   - Level 2: `numFmt=lowerLetter`, `lvlText="(%3)"`
   - Level 3: `numFmt=lowerRoman`, `lvlText="(%4)"`
   - Wire `pPr` indents per level to match the document's existing visual layout.
2. For every numbered paragraph: **strip the manually typed token** from the run text (preserving everything after it, including formatting of remaining runs), attach `w:numPr` with the correct `ilvl` and the new `numId`. Use `startOverride` for legitimate restarts.
3. Map paragraphs to named styles (`LDArticle`, `LDSection`, …) added to `styles.xml`, rather than direct formatting, so firms can restyle globally.

### Cross-reference conversion

- **Detect** textual references in body runs: `Section 4.02`, `Sections 4.02 and 4.05`, `Sections 4.02 through 4.06`, `Article VII`, `clause (b)`, `Section 4.02(b)(i)`, case-insensitive variants. Also handle `this Section N` (self-reference).
- **Resolve** each to a target paragraph in the inferred tree. Unresolvable references (pointing at numbers that don't exist — a huge real-world bug class) go into the plan as warnings: this alone is valuable ("your document references Section 9.4; there is no Section 9.4").
- **Convert**: insert a bookmark (`_LDRef_<n>`) spanning the target paragraph, replace the reference text with a field: `{ REF _LDRef_<n> \r \h }` (`\r` = paragraph number, `\h` = hyperlink). For ranges/lists, convert each number in the phrase to its own field, keep connective words as text.
- Set `w:updateFields` in settings.xml so Word refreshes fields on open.

### Hard edge cases (handle explicitly, with tests)

- **Tracked changes present** (`w:ins`/`w:del` anywhere): v1 refuses with a clear message — "accept or reject all changes first." Do not attempt repair through revision marks.
- **Tables**: numbered paragraphs inside table cells participate in numbering; handle `w:tbl` traversal.
- **Fields already present**: existing REF/TOC fields must survive untouched; existing REF fields pointing at paragraphs we renumber must keep working (they will, since we keep native numbering — verify in tests).
- **Section breaks, headers/footers, footnotes**: out of scope for renumbering but must round-trip byte-safe.
- **Mixed lists**: exhibits/schedules (`Exhibit A`, `Schedule 1`) are separate sequences — detect and keep separate, never merge into the clause hierarchy.
- **Definitions with embedded parentheticals**: `("Confidential Information")` must never be parsed as a list token (tokens are only recognized at paragraph start).
- **Docs that are fine**: if inventory finds consistent native numbering and no anomalies, report "nothing to fix" and change nothing.

---

## Package: `cli`

`legal-down fix input.docx -o output.docx --report plan.json`. Exists to run the corpus, debug, and let power users script. ~50 lines around core.

---

## Package: `web`

Single static page (host on GitHub Pages/Vercel — free):

1. Drag-and-drop a .docx. **File is processed entirely in-browser** — state this prominently, link to the source, and note the page works offline (verifiable proof of the privacy claim; ship it as a PWA so it literally runs with wifi off).
2. Show the repair plan: side-by-side old/new numbering, xref conversions, warnings list (including broken-reference findings).
3. "Download fixed .docx."

SEO matters: the landing copy should target the actual searches ("word numbering messed up legal document", "fix cross references contract word"). This page is the discovery front door.

## Package: `addin`

Office.js taskpane add-in:

- Buttons: **Scan** (renders the plan in the sidebar, including the broken-xref audit) and **Apply**.
- Implementation: get the document via `body.getOoxml()` (or package-level OOXML), run the same core engine, write back via `insertOoxml` replace. Wrap in a single undo-able operation where possible; before Apply, offer "save a backup copy."
- Manifest hosted on the same static host. Provide sideloading instructions for solos (no AppSource dependency to start; submit to AppSource later — expect a review cycle).
- The add-in does **not** need its own logic — it is a thin shell over `core`. Keep it that way.

---

## Testing (build this FIRST)

The corpus is the project's real asset:

- `corpus/broken/*.docx` — 20+ real-world broken documents (manually numbered, half-converted, pasted-in clauses, wrong restarts, dead xrefs). Start by scripting *generation* of broken docs (take a clean agreement, apply mutation operators: retype numbers as text, shuffle a section, break a restart, insert a mis-numbered pasted clause, delete a referenced section) so there's an unlimited supply with known ground truth.
- `corpus/golden/*.plan.json` — expected repair plans, snapshot-tested.
- **Invariant tests on every corpus doc**: (1) plain-text equality modulo number tokens; (2) output opens with zero errors under strict OOXML validation; (3) idempotence — running the tool on its own output produces an empty plan; (4) simulated post-edit: programmatically insert a paragraph into the fixed doc, recompute Word numbering semantics, assert everything renumbers.
- CI runs the corpus on every commit.

**Definition of done for v1:** engine passes the full corpus; web app fixes a real 50-page agreement in-browser in <5 seconds; add-in applies the same fix inside Word; a document repaired by the tool survives subsequent manual editing in Word without breaking.

### Deterministic formatting-repair extension

The same plan/rebuild pipeline also repairs high-confidence formatting damage:
novelty or randomly swapped fonts, inconsistent point sizes, accidental
bold/italic/underline, paragraph alignment and indentation drift, compressed or
excessive spacing, highlighted placeholders, and uneven section margins. It
emits reusable `LDTitle`, `LDSubtitle`, `LDSection`, `LDBody`, and `LDSignature`
styles rather than layering on more direct formatting. Flat agreements numbered
`1.`, `2.`, `3.` are represented as native top-level decimal lists.

This extension keeps the zero-content-loss rule. Run-on paragraphs with no
internal sentence boundary are surfaced as editorial warnings and are never
rewritten automatically.

---

## Milestones

1. **M0 — Corpus + mutation generator + invariant harness.** No engine yet. (This forces the eval to exist before the code.)
2. **M1 — Inventory + inference**: correct plans on the corpus, including the `(i)` ambiguity and anomaly flagging.
3. **M2 — Rebuild**: clean abstractNum + numPr rewrite, invariants green.
4. **M3 — Cross-reference detection/conversion + broken-ref audit.**
5. **M4 — CLI + web app.**
6. **M5 — Word add-in + sideloading docs.**
7. **M6 — Hardening**: tables, exhibits, tracked-changes refusal, idempotence, AppSource submission.

## Non-goals (v1)

- No markdown editor or drafting mode (possible v2, layered on the same engine)
- No AI/LLM anywhere in the pipeline — the engine is deterministic by design and by promise
- No cloud, accounts, storage, or telemetry
- No redline/tracked-changes processing (refuse cleanly)
- No .doc (legacy binary) support — .docx only
