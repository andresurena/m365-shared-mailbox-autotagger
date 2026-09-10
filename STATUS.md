# Status

## What exists and works today

The full pipeline is built and verified working: `worker.js` subscribes to Microsoft
Graph change notifications on a shared mailbox's Sent Items, detects Send-on-Behalf
replies via the `sender`/`from` distinction, tags the reply, and finds/tags the
original message via `conversationId`. A daily Cron Trigger renews the Graph
subscription indefinitely with no manual upkeep. Two commits, both 2026-07-27: the
initial build, then an expansion of the README's supporting evidence (six Reddit
threads spanning 2013–2023, cited to show this is a long-standing, unsolved Microsoft
365 gap, not a one-off complaint).

Known from context (not from the repo's own files): a real, non-genericized deployment
of this exists in production for one client's shared mailbox — this repo is the
sanitized, reusable reference version, with `example.com` / `Alice·Bob·Carol`
placeholders standing in for real values. The repo itself is not what's deployed
anywhere; it's documentation + a template.

## In flight

Nothing — `grep` for TODO/FIXME/XXX across the code returns nothing, and there's been
no commit activity since the two same-day commits. The repo is in a stable, "finished
as documented" state, not mid-feature.

## Open questions

- No automated tests exist anywhere (confirmed by audit). Acceptable for a single-file
  webhook handler with manual verification, or worth adding given it writes to live
  mailboxes on every notification?
- `assets/sent-items-tagged-linkedin.png` exists in the repo but isn't referenced by
  `README.md` (only `thread-tag.png` and `sent-items-tagged.png` are linked). Kept
  deliberately for external reuse (e.g. a promotional post), or orphaned?
- No `wrangler.toml` — deploys are manual/ad hoc via whichever of the three documented
  methods the deployer picks. Fine for a single maintainer; worth codifying if this
  gets deployed more than once.

## Next (best guess — correct me)

- If/when this is deployed for a second mailbox or client, extract the per-deployment
  bits (`CATEGORY_MAP`, mailbox address) more cleanly instead of hand-editing
  `worker.js` per deployment.
- Consider a `wrangler.toml` for reproducible deploys, replacing the current
  "however you deploy Workers" flexibility with one documented path.
- The README's own "Limitations" section flags that only the *reply* gets tagged, not
  the original message directly (they're linked via conversation matching, not the
  same tag applied at the source) — worth deciding whether that's worth solving with
  the more involved Graph-relay approach that was considered but not built.
