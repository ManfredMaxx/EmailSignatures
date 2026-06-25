# Signify — Handoff

_Last updated: 2026-06-25_

## Where we are right now

**v1.2.0 is deployed to a test user (Dan) and v1.3.0 fixes are pushed (code-only).** GitHub Pages is live; the add-in is admin-deployed in M365 and appears in new Outlook on Windows.

Dan field-tested v1.2.0 and found real issues; all are fixed in **v1.3.0** (pushed to `main`, code-only, manifest untouched):
- **Send guard now fails CLOSED (R0 Prime Directive).** `commands.js` rewritten: no silent send on read error; `Office.onReady` init (fixes the New-Outlook hang); ~4 s block-don't-hang safety net; "Send Anyway" override only when a signature is genuinely missing.
- **Editor redesigned** — modern, roomier toolbar; refined spacing.
- **Image/Link by URL fixed** — inline URL input bar (replaced `window.prompt`, which new Outlook silently blocks).
- **Auto-save** — localStorage working copy + restore; nothing lost if Save is forgotten.
- **Colour swatch palette**; **tighter line spacing** (`defaultParagraphSeparator=div`).

These are hosted-file changes → live ~1 min after push; Dan reopens the compose window (or restarts Outlook) to pick them up.

## Pending decision (Dan)

**Manifest `SendMode="Block"` to fully close R0.** Code fails closed, but the *manifest* is still `PromptUser`, which means if the add-in can't load/offline, Outlook still sends (fail-open at the platform level). Only `SendMode="Block"` closes that — but it's a manifest change (≤24 h propagation) and wants an admin offline policy (`OnSendAddinsEnabled` via Exchange PowerShell). Asked Dan to choose strictness (Block+override / Block strict / keep PromptUser). See `04_DECISIONS.md`.

## What's next (in order)

1. **Dan re-tests v1.3.0** in Outlook (reopen compose): the redesign, URL buttons, autosave, and especially the **send guard** (Send with/without a signature — should now show a clear block, not hang).
2. **Resolve the `SendMode="Block"` decision** → if yes, change manifest (+ version bump) and guide the admin offline policy.
3. **Mobile single source (R7)** — start with the roamingSettings-in-mobile-runtime spike (see BACKLOG / DECISIONS).
4. **Easy image hosting** (R6, path B).

## Flagged to revisit

- **R0 full closure** — manifest `SendMode="Block"` + admin offline policy (pending decision above).
- **Easy no-backend image hosting** (R6, path B) — see `02_BACKLOG.md` / `04_DECISIONS.md`.
- **Mobile single source** (R7) — see `02_BACKLOG.md`.
