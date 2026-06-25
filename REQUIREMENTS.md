# Signify — Critical Requirements

_Binding product constraints. **Every recommendation and design decision must be checked against this list.** If an approach can't meet a CRITICAL requirement, say so explicitly and either propose a compliant alternative or surface the trade-off for an explicit decision — never quietly trade one away._

_Last updated: 2026-06-23. Status reflects the v1.2.0 build._

## Critical (non-negotiable)

> ### 🛑 R0 — PRIME DIRECTIVE (outranks everything below)
> **An email must NEVER be sent without a signature unless the user gives explicit, conscious permission** (a deliberate "Send Anyway"). The send guard must **fail CLOSED**: if it cannot confirm a signature is present — because of an error, a timeout, or the add-in being unavailable/offline — it must **block or prompt, never silently send**. A silent or accidental unsigned send is the catastrophic *fail-deadly* outcome, to be prevented above all else, even at the cost of convenience. Every change to a send path is judged against this first.
>
> **Platform facts (verified, Microsoft Smart Alerts docs):** `SendMode="PromptUser"` and `"SoftBlock"` **send the message** if the add-in errors/can't load/is offline (fail-open). Only `SendMode="Block"` refuses to send when the add-in is unavailable (fail-closed). For full coverage when Outlook launches offline, the admin must also set the `OnSendAddinsEnabled` (Web/new Windows) or `OnSendAddinsWaitForLoad` (Mac) mailbox policy. Residual platform gaps that the add-in alone cannot close: Simple-MAPI sends, and offline-launch without the policy.
>
> **Status:** 🟡 Handler hardened to **fail-closed** (v1.3.0: no silent send on read error/timeout; no hang; conscious "Send Anyway" override offered only when a signature is genuinely missing). Full closure (`SendMode="Block"` + admin offline policy) **deferred by Dan (2026-06-25)** as a critical backlog decision — interim `SendMode="PromptUser"` accepted to avoid destabilising live Outlook; a more robust system is ultimately intended. See `02_BACKLOG.md`. A prominent, **off-by-default, warning-gated user toggle** to disable the check ships in v1.4.0 (`sendBlockDisabled`); any future strict mechanism must keep an equivalent toggle.

| # | Requirement | Why | Status |
|---|---|---|---|
| **R0** | **Fail-safe send guard** — never send unsigned without explicit permission; fail **closed** on any uncertainty | **Prime Directive.** The compliance catastrophe is a silent/accidental unsigned send | 🟡 Handler fail-closed (this build); full hardening pending — see callout above |
| **R1** | **Zero recurring cost.** A free replacement for Exclaimer — no per-user SaaS fees, no mandatory paid backend/server. | The entire reason the project exists. | ✅ Met — static files on free GitHub Pages; per-user data in the mailbox. |
| **R2** | **Freeform per-user signatures via a WYSIWYG editor.** Each user designs their own signature visually (fonts, colours, layout, images, compliance text). **NOT** a fixed template they fill in. | Compliance-conscious industry; signatures differ per person and must carry license numbers / regulatory text laid out as required. | ✅ Met — v1.2.0 contenteditable editor. |
| **R3** | **Correct placement.** Signature lands **after the new message text and before the quoted thread** — never dumped at the very bottom. | The core gap in server-side stamping; the reason we chose a client add-in. | ✅ Met on desktop/web via `setSignatureAsync`; same API is available on mobile. |
| **R4** | **Send safety.** The user is warned/blocked if they try to send with no signature, with an emergency override. | Compliance — a missing signature must not go out silently. | ✅ Met on desktop/web (`OnMessageSend`). ⚠️ Not available on mobile (mitigated by auto-insert — see R7). |
| **R5** | **Self-service & easy to deploy.** One-time admin deploy; each user sets up and edits their own signature inside Outlook with no hand-holding. Handable to a non-technical customer without explanation. | It is being given to others to run. | ✅ Met — single manifest upload + README. |
| **R6** | **Images in signatures.** Users can add a logo/image; easy, reliable image hosting still to be solved. | Branded signatures need a logo. | 🟡 Partial — URL + size-guarded paste/drag shipped; easy no-backend hosting flagged (see `02_BACKLOG.md` / `04_DECISIONS.md`). |
| **R7** | **Cross-platform, single source of truth.** Signatures must also work in Outlook **mobile**, and editing the signature in one place updates it **everywhere** — no duplicate or divergent systems. | Stated critical requirement (2026-06-23). | 🔵 Planned — achievable **without a backend** via a mobile `OnNewMessageCompose` handler reading the same `roamingSettings` store (see `04_DECISIONS.md` / `01_ROADMAP.md`). |

## Context / scale
- Microsoft 365 tenant, ~5 users.
- An Outlook shop (desktop, web, mobile) — not targeting non-Outlook mail clients.

## Engineering constraints (the "how", not the "what")
- No third-party runtime dependencies (plain Office.js + the web platform).
- **XML add-in-only manifest** — required for Outlook mobile event-based add-ins; the unified JSON manifest is not yet supported on mobile.
- **Admin-deployed** — required for event-based activation and the send-block.
- Maintain the doc system; update docs after each change/build.

## How to use this doc
Before recommending an approach, confirm it against **R0–R7**. Common violations to watch for:
- **Any send path that fails open** (allows the send when it can't confirm a signature) → breaks **R0**. When in doubt, block/prompt.
- A paid backend or SaaS → breaks **R1**.
- A fixed/AD-template signature in place of freeform → breaks **R2**.
- Bottom-of-email placement (e.g. a transport-rule disclaimer) → breaks **R3**.
- Two independently-edited signature definitions kept "close" → breaks **R7** (single source means *one* store feeding all platforms, not two systems in sync).
