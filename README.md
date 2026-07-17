# Legal Down

Repair broken numbering and cross-references in legal Word documents without
uploading the document anywhere.

Legal Down is a deterministic, client-side OOXML repair engine. It turns typed
clause labels into native Word multilevel numbering and textual references into
live `REF` fields, while preserving document content and unrelated package
parts. The web app works offline; the CLI is useful for batch checks; the Word
add-in is a thin interface over the same engine.

> Early alpha: test on copies of non-confidential documents. A repair is never
> applied when tracked changes are present or structural anomalies have not
> been confirmed.

## Try it locally

Requires Node.js 22+ and pnpm 11+.

```bash
pnpm install
pnpm check
pnpm dev
```

Then open the URL printed by Vite. The CLI can scan before applying:

```bash
pnpm --filter @legal-down/cli start -- scan agreement.docx
pnpm --filter @legal-down/cli start -- fix agreement.docx -o agreement.fixed.docx --report plan.json
```

## Safety model

- `.docx` bytes remain on the device. There is no backend, account, analytics,
  document logging, or document-content telemetry.
- `scan` produces a reviewable plan. `fix` refuses tracked changes and
  unresolved anomalies.
- Untouched ZIP entries are copied without XML reserialization.
- Existing fields and out-of-scope document parts are retained.

See [Architecture](docs/architecture.md), [Word add-in sideloading](docs/addin.md),
and [Contributing](CONTRIBUTING.md).

## Packages

| Package | Purpose |
| --- | --- |
| `@legal-down/core` | DOM-free OOXML inventory, inference, planning, and rebuild |
| `@legal-down/cli` | Node.js scan/fix commands |
| `@legal-down/web` | Offline-capable drag-and-drop PWA |
| `@legal-down/addin` | Office.js Word task pane |

## License

MIT. Legal Down is not a law firm and does not provide legal advice.
