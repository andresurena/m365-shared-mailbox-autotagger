# 000 — Baseline: the system as built

This is spec v1 written after the fact — what the original spec would have said,
derived from the shipped code. Future specs reference this as "against the baseline...".

## Problem

Microsoft 365 shared mailboxes have no native way to show who replied to a message.
The only usable signal is Exchange's "Send on Behalf" delegation mode, and even that
isn't surfaced or tagged anywhere automatically — someone has to remember, or open
every message and check.

## Scope

One background service, watching one shared mailbox, tagging replies with who sent
them. Not a helpdesk tool, not an inbox client, not a general-purpose Graph
integration — deliberately the smallest thing that closes this one gap.

## Mechanism

1. **Detection.** For a Send-on-Behalf message, Microsoft Graph's `sender` field holds
   the real person; `from` holds the shared mailbox. When they differ, someone replied
   on-behalf. When they match (a normal send, or Send As), there's no real sender to
   attribute — skip it.
2. **Trigger.** A Graph subscription (change notification / webhook) on the shared
   mailbox's Sent Items folder calls the Worker the instant a new message lands there.
   No polling.
3. **Attribution.** The sender's display name is looked up in a hardcoded
   `CATEGORY_MAP` and matched to an Outlook category already created in the shared
   mailbox.
4. **Tagging.** The reply itself gets the category via a Graph `PATCH` on its
   `categories` property (append, not overwrite — existing categories are preserved).
   The original message is found by matching `conversationId` across the mailbox
   (Graph doesn't support filtering on `id`, so the current message is excluded
   client-side afterward) and tagged the same way.
5. **Upkeep.** Graph subscriptions expire after 7 days max with no auto-renew. A daily
   Cron Trigger calls the same subscription-creation logic, which is idempotent —
   renews if a subscription already exists, creates one if not.

## Access model

The Graph app registration uses application (not delegated) permissions —
`Mail.ReadWrite` — authenticated via client-credentials flow, no signed-in user. That
permission is tenant-wide by default, so it's deliberately narrowed with an Exchange
Application Access Policy scoping the app to the one target mailbox only. Verified via
`Test-ApplicationAccessPolicy`, not assumed.

## Explicit non-goals

- Not a database of who-replied-to-what — the tag is the record; there's no separate
  queryable log.
- Not real-time-perfect: relies on staff correctly selecting the shared mailbox in the
  From field. A reply sent from someone's personal address is invisible to this system,
  same as it is today without it.
- Not multi-tenant or multi-mailbox out of the box — `MAILBOX` and `CATEGORY_MAP` are
  single-deployment configuration, hand-edited per install.

## Interfaces this system depends on

- Microsoft Graph v1.0 (`/subscriptions`, `/users/{mailbox}/messages`) — no other
  Microsoft API surface.
- Cloudflare Workers runtime (`fetch` + `scheduled` handlers) — the implementation is
  Workers-specific, but the mechanism itself is host-agnostic (any webhook-capable
  runtime would work, per the README).

## What "done" looked like for v1

Verified end-to-end: a reply sent on-behalf from the shared mailbox gets tagged within
seconds, and the original message it replied to gets the same tag via conversation
matching. That's the whole baseline — everything past this point is a new spec.
