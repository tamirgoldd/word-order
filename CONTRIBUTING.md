# Contributing

Word Order treats document integrity as the primary feature. Before opening a
pull request:

1. Add a synthetic fixture that demonstrates the document pattern.
2. Add or update the expected repair plan.
3. Run `pnpm check`.
4. Confirm the input fixture contains no client or personal information.

Core changes should preserve untouched OOXML parts byte-for-byte and must fail
closed when structure is ambiguous. New runtime dependencies need a clear
justification.
