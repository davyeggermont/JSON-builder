---
name: tidy-memory
description: Clean up the project memory. Reads memory/*.md and CLAUDE.md, merges duplicates, drops stale entries, tightens wording, and flags entries that have become standing rules so the user can move them into CLAUDE.md. Use when the user runs /tidy-memory or asks to clean up, prune or consolidate project memory.
---

# Tidy project memory

Clean up `memory/decisions.md`, `memory/open-questions.md` and `memory/conventions.md` against
each other, against `CLAUDE.md`, and against the current repo. Edit the memory files directly.
**Never edit `CLAUDE.md`.** Only propose changes to it.

## 1. Read everything first
- `CLAUDE.md`, all three files in `memory/`, and the format comment at the top of each.
- `git log --oneline -20` and `git status`, to see what has changed since the entries were written.
- When an entry names a file, column, sheet, function or behavior, check that it still exists
  (Grep/Glob, or read `README.md`, `json-builder.html` and `tools/`).

## 2. Decide what happens to each entry
Go through every entry and give it exactly one outcome:

| Outcome | When |
|---------|------|
| **Merge** | Two or more entries say the same thing, within a file or across files. Keep one, under the most fitting file, with the earliest date. Merge their reasons. |
| **Drop** | It's stale: it refers to something that no longer exists; it's a decision marked superseded whose replacement fully covers it; it's an open question the code, git history or another entry now answers; or it duplicates something `CLAUDE.md` already says. |
| **Move** | It's in the wrong file. For example, an answered question belongs in decisions, or a decision that is really a standing preference belongs in conventions. |
| **Tighten** | It's kept but wordy. Cut it to one fact, keep the *why*, remove session narration and hedges. Don't change the meaning. |
| **Flag for CLAUDE.md** | It has become a rule: it applies to every session, is stated as always/never, has held for a while (several entries or repeated requests), and is short enough for one line in CLAUDE.md. Leave it in memory; the user decides. |
| **Keep** | Already fine. |

Rules for deciding:
- If you aren't sure whether something is stale, keep it and list it under "Unsure" in the report. Don't drop it.
- Entries marked "confirm with the user" or "from README" may be flagged as unconfirmed. Don't drop them just for that.
- Keep absolute dates. Never invent a date, a reason or a decision that isn't in the files or the repo.
- Keep each file's format and its header comment exactly as they are. Keep decisions newest-first.
- If a file ends up empty, leave the header, the format comment and `_None yet._`.

## 3. Apply the edits
Edit the three memory files to carry out Merge, Drop, Move and Tighten. Don't commit.

## 4. Report
Reply with:

1. **Changes:** a short list grouped as merged, dropped (with the reason each was stale), moved and tightened. Count the kept entries rather than listing them.
2. **Unsure:** entries you left alone but suspect are stale, each with the question that would settle it.
3. **Proposed for CLAUDE.md:** for each flagged entry, the exact line to add and the section it belongs under, ready to paste. Say whether to delete the memory entry once it's moved.
4. One line telling the user to review the edits with `git diff memory/`.
