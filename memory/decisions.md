# Decisions

Choices we made and why. Newest at the top. Never edit an old entry's reasoning; if a
decision is reversed, add a new entry and mark the old one `**Superseded by:** <date title>`.

<!-- Format:
## YYYY-MM-DD: Short title
**Decision:** what we chose, in one or two sentences.
**Why:** the reason, including the constraint or problem that drove it.
**Rejected:** alternatives we considered and why not (optional).
-->

## 2026-10-21: Single offline HTML file, no install
**Decision:** The tool is one self-contained `json-builder.html` that runs in Edge/Chrome from disk.
**Why:** Users must be able to copy two files to another PC and use it with no install, no
internet, and no data leaving the machine. (Recorded from README; reasoning not discussed in session.)

## 2026-10-21: Find Excel columns by header name, not position
**Decision:** The builder looks up columns by their row-1 header; column order doesn't matter.
**Why:** Users edit the Excel themselves and may reorder or add columns (e.g. "Datum 9", new codes).
(Recorded from README.)
