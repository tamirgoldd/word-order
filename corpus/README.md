# Synthetic corpus

The committed corpus is declarative so it is reviewable and cannot contain
client material. Run `pnpm corpus:generate` to create 24 deterministic `.docx`
files under `corpus/broken/`; generated binaries are gitignored.

Each scenario begins with a small clean agreement and applies named mutation
operators such as typed-number conversion, sequence jumps, deleted targets,
table insertion, and tracked-change wrapping. The matching golden JSON captures
the stable part of the expected plan. Contributors should add a minimized
synthetic scenario for every production bug.
