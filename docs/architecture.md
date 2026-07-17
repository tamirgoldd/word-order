# Architecture

Legal Down uses one deterministic engine across three interfaces. The core
opens the OOXML ZIP, inventories paragraphs, infers hierarchy, creates a repair
plan, and only then rebuilds the changed package parts.

The engine does not use browser DOM APIs. Reads are tolerant, while writes are
targeted string transformations so unknown XML and untouched ZIP members are
not normalized accidentally. A plan records every numbering and reference
decision and has a digest tied to the source document.

## Failure boundaries

- Tracked changes block planning and repair.
- Missing or malformed mandatory OOXML parts produce a typed error.
- Ambiguous sequences produce confirmation-required anomalies.
- A clean document returns its original bytes without rezipping.
- Cross-references that cannot be resolved are reported and left as text.

The corpus generator turns declarative agreements into `.docx` fixtures and
applies known mutation operators. Tests assert plan snapshots, preserved plain
text, untouched package parts, idempotence, and valid package relationships.
