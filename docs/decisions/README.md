# Decisions

Architectural decision records (ADRs). One file per decision; each captures *why* a structural choice was made so future contributors can reconstruct (or revisit) the reasoning without trawling commit history.

## Format

Each decision file follows this skeleton:

- **Status** &mdash; Accepted / Superseded / Rejected, plus the codification date.
- **Context** &mdash; the situation that forced a choice; the alternatives considered.
- **Decision** &mdash; the choice made, stated plainly.
- **Why** &mdash; the reasoning. The most important section.
- **Consequences** &mdash; what this commits us to, both operationally and strategically.
- **Cross-references** &mdash; related docs, code, issues.
- **Reconsider when** &mdash; concrete signals that would make us re-open the decision.

Decisions are immutable once accepted. To change one, write a new decision file that supersedes it and update the old file's status.

## Index

- [tests-need-visio.md](tests-need-visio.md) &mdash; The test suite drives a real Visio instance via COM; no mock layer.
