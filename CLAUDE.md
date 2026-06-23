# CLAUDE.md — Signify (EmailSignatures)

Project guidance for Claude Code. Auto-loaded every session; update when architecture or conventions change.

## What this is

A Microsoft 365 Outlook add-in called **Signify**. Each user designs their own email signature via a WYSIWYG rich-text editor embedded in Outlook's task pane, inserts it with one click before sending, and cannot accidentally send without it — blocked by an `OnMessageSend` handler.

No backend server. Five static files hosted on GitHub Pages, deployed org-wide via one manifest upload to M365 Admin Center.

**Audience:** ~5 users in a compliance-conscious industry. Every signature must include license numbers, regulatory disclaimers, and other per-person compliance text.

## File structure

```
EmailSignatures/
  manifest.xml              ← uploaded to M365 Admin Center to install the add-in
  taskpane/
    index.html              ← WYSIWYG editor UI
    index.js                ← editor logic: load, save, insert
  commands/
    commands.html           ← blank entry point for OnMessageSend runtime
    commands.js             ← send-block: checks for sig-marker, blocks if absent
  assets/
    icon-16.png / icon-32.png / icon-80.png
    generate-icons.html     ← browser-runnable canvas tool to regenerate icons
  README.md                 ← customer-facing guide: admin setup + user guide
```

## Hosting

GitHub Pages at `https://ManfredMaxx.github.io/EmailSignatures/`

Manifest GUID: `DCC79AB3-8F0C-4C28-90EF-E27E06CD08DF`

## Key technical facts

- **Signature storage:** `Office.context.roamingSettings` — per-user, stored in the user's Exchange mailbox, syncs automatically to all their Outlook clients. No backend.
- **Signature insertion:** `item.body.setSignatureAsync()` (Mailbox 1.10) — places signature after new text, before quoted thread. Falls back to `prependAsync` if unavailable.
- **Send block:** `OnMessageSend` (Mailbox 1.12) with `SendMode="PromptUser"`. Checks for a hidden `<div id="sig-marker">` in the body. User sees error + can override in emergencies.
- **Minimum requirement:** Mailbox 1.12. Supported by M365 subscription Outlook on recent builds.
- **HTML compatibility:** Signature HTML must use table-based layout and inline CSS only. No flexbox, no grid, no CSS classes, no external fonts. Web-safe fonts only (Arial, Georgia, Verdana, Calibri, Trebuchet MS, Times New Roman).

## Editor approach

WYSIWYG via `contenteditable` div with a custom toolbar. No third-party editor library — Quill/TinyMCE output CSS classes that get stripped by email clients. `document.execCommand` produces inline-styled HTML natively.

**Image support:** URL-based insertion (recommended for email delivery) + paste/drag as base64 (convenience; fine for Outlook compose but adds email weight). The sig-marker must survive all code paths that call `setSignatureAsync`.

## Deployment

Push changes to `main` → GitHub Pages serves immediately. M365 picks up manifest changes within 24h; Outlook may need a restart to reflect them.

## Doc Map

| File | Job |
|---|---|
| `CLAUDE.md` | This file — dev conventions + doc map |
| `README.md` | Customer-facing admin setup guide + user guide |
| `00_HANDOFF.md` | Current state snapshot — where are we right now? |
| `01_ROADMAP.md` | Direction: Now / Next / Later / Gated |
| `02_BACKLOG.md` | Open tasks only (no done items) |
| `03_CHANGELOG.md` | Shipped history, versioned, append-only |
| `04_DECISIONS.md` | Dated log of decisions, assumptions, and resolved unknowns |

**Standing rule:** Update `00_HANDOFF.md` at the end of every session. When shipping, remove tasks from `02_BACKLOG.md`, append a line to `03_CHANGELOG.md`, and bump the version in `manifest.xml`.
