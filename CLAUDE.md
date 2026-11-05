# CLAUDE.md

## Project memory (`memory/`)

Read all three files in `memory/` at the start of a session. They're short; keep them that way.
Write to them as things happen, not at the end. Each file has its format in an HTML comment at the top.

| File | Write when | Format |
|------|------------|--------|
| `memory/decisions.md` | We settle a non-obvious choice (design, data format, tooling, trade-off) and the reason would be lost otherwise. Not for things obvious from the code or git log. | `## YYYY-MM-DD: Title` + **Decision:** / **Why:** / optional **Rejected:**. Newest on top. To reverse one, add a new entry and mark the old one **Superseded by:**. |
| `memory/open-questions.md` | Something is unclear, blocked on the user, or deferred ("let's decide later"). | `- **Question?** Context. (added YYYY-MM-DD)`. When answered, delete it and record the outcome in `decisions.md` or `conventions.md`. |
| `memory/conventions.md` | The user asks for the same thing a second time, or says "always"/"never"/"from now on". | Bullet under an area heading, phrased as an instruction, with *Why:* if given. Edit in place when a rule changes. |

Rules:
- Use absolute dates (today's date), never "yesterday" or "last week".
- One fact per entry; no session narration or summaries of what was done.
- If something you're about to write contradicts an existing entry, update or supersede that entry instead of adding a duplicate.
- Mention in your reply when you've written to memory, in one line.
