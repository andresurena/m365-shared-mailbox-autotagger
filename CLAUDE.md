# m365-shared-mailbox-autotagger

A single Cloudflare Worker that watches a Microsoft 365 shared mailbox and automatically tags every reply (and the message it replied to) with who actually sent it — using Send on Behalf's sender/from distinction and Microsoft Graph change notifications, no third-party SaaS.

## Stack

- Plain JavaScript, ES modules, zero dependencies — no `package.json`, no bundler, no build step. Everything lives in one file, `worker.js`.
- Runtime: Cloudflare Workers (`fetch` handler for the webhook + `scheduled` handler for the daily Cron Trigger that renews the Graph subscription).
- No `wrangler.toml`/`wrangler.jsonc` — the README documents three interchangeable deploy paths (dashboard paste, `wrangler deploy`, or the Cloudflare API), none enforced by config in-repo.
- External integration: Microsoft Graph API v1.0, authenticated via OAuth2 client-credentials flow against Entra ID (app-only, no delegated user).
- Config/secrets: Cloudflare Worker environment variables and encrypted Secrets, set via dashboard or API — no `.env` file in the repo.
- No test framework, no CI (no `.github/` workflows). Verification is manual, documented as a procedure in the README ("Test it" — send a throwaway email, check the tag appears).

## Conventions

The patterns the code already follows — match these, don't impose new ones:

- `export default { fetch, scheduled }` — the standard Workers module shape.
- String concatenation with `+`, not template literals, throughout.
- Null-safety via explicit `&&` chains (`msg.sender && msg.sender.emailAddress && ...`), not optional chaining (`?.`).
- `UPPER_SNAKE_CASE` for the module-level config object (`CATEGORY_MAP`) and all `env.*` variable names; `camelCase` for functions and locals.
- Liberal `console.log`/`console.error` — this is the primary observability mechanism (Cloudflare's Logs panel is the only introspection available), not incidental debug noise. Keep logging generous in new code here.
- Comments explain *why* (Graph API quirks, expiry limits, non-obvious constraints), never *what* the code visibly does.
- Documentation is treated as first-class: the README carries a full setup walkthrough, a "Gotchas" section cataloguing real debugging incidents hit while building this, and an honest "Limitations" section. New features should extend that same standard, not just add a one-line mention.

## Decisions

- 2026-07-27 — Cloudflare Workers chosen as the hosting layer over Azure Functions/AWS Lambda — reused infrastructure already trusted for this engagement rather than introducing a new vendor/trust boundary (stated in README).
- 2026-07-27 — Graph app permission (`Mail.ReadWrite`) restricted to a single mailbox via an Exchange Application Access Policy, rather than left at its default tenant-wide scope — least privilege for a single-purpose background service (stated in README).
- 2026-07-27 — Requires **Send on Behalf**, not Send As — the only delegation mode where Graph exposes `sender` separately from `from`, which the whole detection mechanism depends on (stated in README).
- 2026-07-27 — No build tooling / no `package.json` — *inferred*: a single-file, zero-dependency Worker doesn't need one; kept deliberately minimal.
- 2026-07-27 — Published as a public, genericized repo (MIT, `example.com` placeholders, no real tenant details) rather than kept as private client work — *inferred* from the README's own framing: written to be reused by anyone hitting the same Microsoft 365 gap, not just the original client.

## Standing instructions

These run automatically. Do not wait to be asked.

1. **STATUS.md — end of every session.** Before finishing any working session,
   update `STATUS.md`: what changed, decisions made (mirror durable ones into
   the Decisions section above), open questions, what's next. Keep it under a
   page; it is a state file, not a log.

2. **OBSERVER.md — every session.** Keep a file called `OBSERVER.md` at the
   project root (ensure it is gitignored — it's about Andres, not the product).
   It is a dated observation journal about how he works, written as a fly on
   the wall. Each entry: (1) what you observed, quoting the actual evidence — a
   line from a prompt, a commit, a doc, a decision; (2) the pattern it
   suggests; (3) one constructive push-back — never soften a pattern that
   costs him money, time or focus; (4) one open question you're still forming
   about him. Rules: append, don't rewrite — the point is watching the read
   change over time. Observations before judgments. Never flatter without a
   footnote of evidence. Read between the lines, but quote the lines you read
   between. Be curious and warm; write like someone who wants to understand
   him, not evaluate him. If you notice the same pattern across projects, say
   so and name the projects.

3. **Decisions are recorded when made.** When a durable decision happens
   mid-session — a convention adopted, a scope boundary drawn, a reversal —
   add it to Decisions above in the same session, dated, with its reason.
   Reversals are recorded with their cause, never silently applied.

4. **Specs are the unit of work.** New work arrives as `docs/specs/*.md` files
   from planning chats. Read the spec, confirm your read in 5 lines, flag
   anything underspecified, then build. If reality forces a deviation from the
   spec, note it in the spec file itself under a `## Deviations` heading.
