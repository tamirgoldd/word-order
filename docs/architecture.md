# Architecture

Legal Down uses one deterministic engine across three interfaces. The core
opens the OOXML ZIP, inventories paragraphs, infers hierarchy, audits severe
formatting drift, creates a repair plan, and only then rebuilds the changed
package parts.

The engine does not use browser DOM APIs. Reads are tolerant, while writes are
targeted string transformations so unknown XML and untouched ZIP members are
not normalized accidentally. A plan records every numbering and reference
decision and has a digest tied to the source document.

Formatting normalization is deliberately thresholded. Ordinary documents are
left alone; normalization activates when the engine sees strong evidence such
as three or more direct font families, four or more direct sizes, novelty fonts,
highlighting, deep indents, or materially uneven margins. The repair maps title,
subtitle, heading, body, and signature paragraphs to `LD*` Word styles and
removes only conflicting direct formatting. Meaningful all-caps emphasis and
bracketed placeholders remain bold, but placeholder highlighting is removed.
Long run-on paragraphs are warnings because changing their wording would violate
the text-preservation invariant.

## Failure boundaries

- Tracked changes block planning and repair.
- Missing or malformed mandatory OOXML parts produce a typed error.
- Ambiguous sequences produce confirmation-required anomalies.
- A clean document returns its original bytes without rezipping.
- Cross-references that cannot be resolved are reported and left as text.
- Formatting cleanup never changes sentence text or silently invents paragraph
  breaks.

The corpus generator turns declarative agreements into `.docx` fixtures and
applies known mutation operators. Tests assert plan snapshots, preserved plain
text, untouched package parts, idempotence, and valid package relationships.
