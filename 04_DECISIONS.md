# Signify — Decisions & Assumptions

_Append-only. Each entry: what was decided, why, and consequences._

---

### 2026-06-22 — Office.js add-in over Exchange transport rules

**Decision:** Build a per-user Office.js task pane add-in rather than a server-side Exchange transport rule that appends signatures centrally.

**Why:** Transport rules inject a single admin-controlled template; they can't accommodate per-user customization (different license numbers, individual contact details, personal branding). The compliance requirement is per-person.

**Consequences:** Requires M365 subscription (not personal/family plans), Mailbox 1.12 for the send-block, and admin deployment via M365 Admin Center. Users must insert their own signature; it is not automatic on every compose.

---

### 2026-06-22 — XML manifest over JSON unified manifest

**Decision:** Use the legacy XML manifest format rather than the newer JSON unified manifest.

**Why:** JSON unified manifest had narrower client support as of mid-2026. XML works across classic Outlook desktop, OWA, and New Outlook universally.

**Consequence:** Will need to migrate to JSON manifest when Microsoft eventually deprecates XML — not imminent.

---

### 2026-06-22 — GitHub Pages for static hosting

**Decision:** Host all add-in files on GitHub Pages (free).

**Why:** No backend needed; `roamingSettings` handles all storage. Free tier covers this use case entirely.

**Consequence:** Repo must be public for free GitHub Pages. No sensitive data in the repo (no user signatures are stored there, no credentials). Manifest GUID is public but is not a secret.

---

### 2026-06-22 — `roamingSettings` for per-user storage

**Decision:** Store each user's signature HTML in `Office.context.roamingSettings`.

**Why:** Stored in the user's own Exchange mailbox; syncs automatically to all their Outlook clients; no backend server, database, or API keys required.

**Consequence:** 32KB limit per key (ample for signature HTML). Admin cannot centrally view or manage individual signatures. If a user loses their M365 account they lose their saved signature.

---

### 2026-06-22 — `SendMode="PromptUser"` (not `"Block"`)

**Decision:** Allow send-block override rather than a hard block.

**Why:** Hard blocking creates support burden during genuine emergencies when a user cannot get their signature working. The compliance goal is reminder + friction, not absolute enforcement.

**Consequence:** Users CAN send without a signature if they explicitly choose "Send Anyway." Acceptable for the stated use case.

---

### 2026-06-22 — WYSIWYG editor over form-based template

**Decision:** Replace the initial form-based signature builder (fixed layout template) with a full WYSIWYG rich-text editor.

**Why:** Compliance-conscious users need fully custom signature layouts — different license number placements, company logos, multi-line formatting, regulatory text in specific positions. A fixed template cannot accommodate this variation.

**Consequence:** Editor must produce inline-styled HTML (no CSS classes) to survive email client rendering. `contenteditable` + `document.execCommand` produces compatible output natively; class-based editors (Quill, TinyMCE) would require post-processing to inline all styles.

---

### 2026-06-22 — Image handling: URL-first with base64 drag/paste support

**Decision:** Primary image insertion is URL-based (user provides a link to a hosted image). Drag-and-drop and clipboard paste are also supported, inserting images as base64 data URIs.

**Why:** URL-hosted images produce smaller emails and can be updated centrally (e.g. logo change). Base64 is convenient for one-off inserts and works well in most modern email clients, but adds ~30% size overhead and may be filtered by aggressive corporate email security.

**Consequence:** Users who want reliable cross-client logo display should host the image (GitHub Pages works for this) and paste the URL. Dragged/pasted images are convenient but may bulk up email threads.

---

### 2026-06-22 — Image research: ship URL-first + guarded paste/drag (path A), flag easy hosting (path B)

**Context:** Before building image support, researched how Exclaimer/CodeTwo handle signature images. base64 data URIs (our paste/drag mechanism) are weak on two fronts: the ~32 KB `roamingSettings` cap **and** deliverability — Gmail strips them, Outlook demotes them to attachments, antispam may quarantine them, and they bloat every reply. The industry standard is hosted images by URL. The "embed that works" the pros use is CID hidden attachments, which a client-side `setSignatureAsync` add-in **cannot** produce (no MIME control at send time). The 32 KB workarounds (Office `CustomProperties`, Graph extended properties, external blob store) add disproportionate complexity — Graph auth or a backend — for a 5-user no-backend tool.

**Decision (with user):** Ship **path A** now — **Insert image → URL** is the recommended path; paste/drag stays for small images, size-guarded (~20 KB/image, ~30 KB total), with a one-time deliverability warning. Keep **path B — easy no-backend image hosting** as the durable follow-up (tracked in `02_BACKLOG.md` / `01_ROADMAP.md`).

**Consequence:** Logos display most reliably when hosted and inserted by URL. Refines the prior image-handling entry above. Sources: Exclaimer & CodeTwo image KBs, mail-signatures.com, Suped, Microsoft Learn (roamingSettings limits / error 9057).

---

### 2026-06-22 — `document.execCommand` accepted despite deprecation

**Decision:** Build the rich-text editor on `contenteditable` + `document.execCommand` (with `styleWithCSS`), no third-party editor library.

**Why:** execCommand is deprecated but supported in every current browser and the Outlook WebView, and it natively emits inline-styled HTML — exactly what email clients need. Editor libraries (Quill, TinyMCE) emit CSS classes that email clients strip, and would violate the no-runtime-dependencies convention.

**Consequence:** If a future browser drops execCommand, the toolbar's formatting would need reimplementing on the Selection/Range API. Output is verified clean (no `<font>`/`class`); font-size uses a size-7→inline-px-span conversion to avoid `<font>` tags. `window.prompt` is used for image/link URLs — noted as a possible reliability follow-up in BACKLOG.

---

### 2026-06-22 — Storage format: `sigFields` object → `sigHtml` string

**Decision:** Store the signature as a raw HTML string under roamingSettings key `sigHtml`, replacing the v1.0/v1.1 `sigFields` form-data object.

**Why:** The WYSIWYG editor's content *is* HTML; storing it directly round-trips losslessly back into the editor. There are no fixed fields anymore.

**Consequence:** On load, if `sigHtml` is absent but a legacy `sigFields` exists, it is rendered once into the editor (one-time migration) and dropped on the next Save. Low stakes — nothing was deployed on the old format yet.

---

### 2026-06-23 — Cross-platform single source (R7): mobile via event-based add-in, NO backend

**Context:** New critical requirement — signatures must also work on Outlook mobile from a single source of truth. Researched current Microsoft docs (May–June 2026).

**Findings (verified against learn.microsoft.com):**
- **Outlook mobile supports event-based add-ins.** `OnNewMessageCompose` fires on mobile for new **and reply/reply-all/forward** compose (not draft edits, not the iOS Share sheet). Microsoft's own sample for this event is "add a signature to a new message."
- **`body.setSignatureAsync` is supported on mobile** (explicitly enabled in compose mode despite mobile's 1.5 baseline) — so mobile gets the *same* after-new-text/before-quote placement (R3) as desktop.
- Mobile add-ins **cannot** show a task pane in compose and have **no `OnMessageSend`** — so no editor and no send-block on mobile. The mobile handler is event-driven (auto-insert), which is the correct pattern there.
- **No Microsoft Graph API for signatures** (they aren't a Graph-managed mailbox setting) — rules out a Graph-based central manager.
- **Outlook roaming signatures do not sync to mobile** — that native feature doesn't help.
- **Exchange Online transport-rule disclaimers DO support `%%attribute%%` tokens** (DisplayName, Title, Department, Phone, …) from Entra ID — but that is an AD-driven *template*, not the freeform WYSIWYG signature, so it can't be the single source for R2. Placement is top/bottom only (bottom on replies) → can't meet R3. Useful only as a fallback or for non-Outlook clients. Dedup across a thread is done via an exception that matches unique disclaimer text (our `sig-marker` is exactly this).

**Decision:** Pursue **single source without a backend**: keep `roamingSettings['sigHtml']` as the one store; add a `<MobileFormFactor>` + `OnNewMessageCompose` handler that reads it and calls `setSignatureAsync`. Desktop/web keep the editor + manual insert + send-block; mobile auto-inserts the same stored signature. Editing happens on desktop/web; mobile stays in sync automatically. This **corrects** the earlier exploratory claim that true cross-platform sync would require a backend.

**Open question to resolve in a spike before building:** confirm `Office.context.roamingSettings` is readable inside the **mobile event runtime** (it's a 1.1 API under mobile's 1.5 baseline, so expected to work, but Microsoft's sample reads `from.getAsync` instead). If it isn't, fall back to `sessionData` or an AD-token transport rule for mobile (accepting the R3 placement compromise there). Sources: Microsoft Learn — *Implement event-based activation in Outlook mobile add-ins*, *Outlook JS APIs supported on mobile*, *Organization-wide disclaimers in Exchange Online*, *Get/Update user mailboxSettings (Graph)*.

---

### 2026-06-25 — R0 Prime Directive: send guard must fail CLOSED

**Context:** Dan elevated "never send unsigned without express permission" to the top requirement (R0), defining the *fail-deadly* outcome as a silent/accidental unsigned send. Researched Smart Alerts failure semantics (Microsoft docs).

**Verified platform facts:** When the OnMessageSend add-in errors / can't load / is offline — `SendMode="PromptUser"` and `"SoftBlock"` **send the message** (fail-open); only `SendMode="Block"` refuses to send (fail-closed). Timeout threshold is ~5 s (long-running dialog) / 5 min (hard timeout). For the *Outlook-launched-offline* gap, the admin must also set `OnSendAddinsEnabled` (Web/new Windows, via Exchange Online PowerShell) or `OnSendAddinsWaitForLoad` (Mac). Runtime **send-mode override** (requirement set 1.14) lets the handler downgrade Block→PromptUser per-event. Residual gaps the add-in alone can't close: Simple-MAPI sends; offline-launch without the policy.

**Decisions:**
- **Code (shipped now, fail-closed):** the old fail-open path (`allowEvent:true` on read error) is **removed**. `commands.js` now: inits `Office.onReady`; on read error/timeout → `allowEvent:false` (block); on missing signature → `allowEvent:false` + `sendModeOverride:PromptUser` (conscious "Send Anyway" = express permission); ~4 s safety timeout that blocks rather than hangs. This also fixes the New-Outlook hang (handler now always completes, and quickly).
- **Manifest `SendMode="Block"` (pending Dan's decision):** the current manifest stays `PromptUser` so the code fix ships instantly (hosted files, ~1 min) without the 24 h manifest-propagation delay. Flipping to `Block` is the only way to close the *add-in-unavailable* fail-open hole; it's a manifest change (≤24 h) and benefits from the admin offline policy. Tracked as a Now item in `01_ROADMAP.md` / `02_BACKLOG.md`.

**Why split the ship:** the code hardening is the high-value, instant part and is safe under either send mode; the manifest flip is slower (24 h) and carries an ops choice (strictness + admin policy), so it waits for an explicit decision. Sources: Microsoft Learn — *Handle OnMessageSend … with Smart Alerts* (send-mode options, add-in-unavailable, timeout, offline behaviour).

---

### 2026-06-25 — Image hosting is a near-term build issue (32 KB will bite often)

**Question (Dan):** can a signature image URL be a local file, or must it be hosted? Will we keep hitting the 32 KB cap?

**Answer / facts:** A signature image **cannot** be a local file. The email is delivered to *recipients* who can't access the sender's disk, and mail clients block `file://` for security. A signature image must be either (a) **hosted at a public http/https URL**, or (b) **embedded** (base64 data-URI or CID attachment). The ~32 KB `roamingSettings` cap only constrains *embedded* (base64) images — those are stored in the mailbox with the signature. **Hosted URLs are a few bytes in storage and dodge the cap**, and also render far better (Gmail strips base64; Outlook demotes it to an attachment). So: yes, embedded images will hit 32 KB frequently — the durable answer is hosting, not embedding.

**Decision / flag:** elevate **R6 "easy no-backend image hosting"** from a nice-to-have to a likely **near-term build** — a small interface so a user's logo gets a hosted URL without manual git/hosting work. Since we already run on GitHub Pages, the cheapest path is logos living in the repo (`assets/logos/…`) referenced by their Pages URL; the open problem is a low-friction way to get a file in (non-technical users can't git-commit — likely an admin one-time drop or a guided upload). A backend image host is the fallback if no-backend proves too clunky. Tracked in `01_ROADMAP.md` / `02_BACKLOG.md`.

---

### 2026-06-25 — Mobile build decisions + de-risk spike

**Decisions (Dan):**
1. Mobile auto-inserts a **dedicated mobile default** signature (not merely the desktop "active" one).
2. Desktop gets an **auto-insert-on-compose option** (alongside today's manual button + send-block).
3. **Separate signatures for New Email vs Reply/Forward** on desktop (like Outlook's native new-vs-reply signatures). Porting the new/reply distinction to mobile is desirable, not critical.

**Implication for the data model:** the signature library gains **role assignments** per signature — e.g. desktop-new, desktop-reply, mobile-new, mobile-reply. Good news for #3 on mobile: `getComposeTypeAsync` **is supported on Outlook mobile**, so the mobile handler can tell new vs reply/forward and pick the matching default — so the new/reply distinction *can* port to mobile.

**Spike (built, awaiting Dan's on-device test):** to avoid risking the live desktop send-block, the de-risk is an **isolated test add-in** in `spike/` (own GUID `8FD29D48-…`), not a change to production. A desktop taskpane writes a test value to `roamingSettings`; an `OnNewMessageCompose` handler reads it back on iOS and reports — via an in-Outlook notification + inserted text — whether (a) the iOS event runtime can read `roamingSettings` and (b) `setSignatureAsync` works on iOS. Manifest min 1.5 (mobile baseline), validated. Dan deploys it alongside Signify, tests on his iPhone, then removes it. Outcome gates the real mobile build.

---

### 2026-06-25 — Spike PASSED: mobile single source confirmed (R7 green-lit)

On-device iPhone test succeeded. The `OnNewMessageCompose` handler read the desktop-written `roamingSettings` value ("desktop-write-OK") **and** `setSignatureAsync` inserted on iOS. Confirms: (1) the iOS event runtime can read `roamingSettings`; (2) desktop→mobile `roamingSettings` sync works; (3) `setSignatureAsync` works on iOS. **The clean, no-backend, single-source mobile path (R7) is open** — proceed to the real mobile build (mobile default(s), new-vs-reply via `getComposeTypeAsync`, desktop auto-insert option). The 4×/day reminder task was disabled. The throwaway spike add-in can be removed from M365.

---

### 2026-06-25 — Feasibility: a desktop program feeding the add-in

**Question (Dan):** could a separate desktop program create/deploy signatures that the add-in consumes, for more flexibility in how they're saved/deployed?

**Verified present data (2026):** (a) **No Microsoft Graph API for Outlook signatures** — still unsupported (Microsoft Q&A, 2026). (b) The add-in's `roamingSettings` are stored as a single MAPI extended property `cecp-<app-guid>` (property set `{00020329-0000-0000-C000-000000000046}`) holding the settings as JSON — **readable** (and, unsupported, **writable**) externally via Microsoft Graph / EWS extended properties.

**Core constraint:** the add-in and a desktop program can't talk directly; they bridge via a shared store. Options scoped:
1. **Inject into `roamingSettings`** via the `cecp` extended property (Graph/EWS) — technically possible, but writes an undocumented internal format (fragile/unsupported) and needs an Entra app registration + mailbox permissions + admin consent. Advanced; not recommended as the primary path.
2. **Add-in FETCHES from a shared store** the desktop program publishes to (per-user, keyed by email) — clean and supported (web + CORS). Cost = a reachable store (free static hosting with privacy caveats, or a small backend). This is the centralized-management model.
3. **Import/export** — desktop designer exports HTML; the add-in gains an Import feature. Easiest, free, but manual hand-off.
4. **Bypass the add-in** — a desktop tool (e.g. the open-source Set-OutlookSignatures, or CodeTwo) deploys the **native** Outlook signature, but loses the send-block + placement that make Signify valuable. _(Set-OutlookSignatures specifics from established knowledge; a search outage prevented re-verification.)_

**Recommendation by goal:** richer *design* → #3; central *deployment* → #2. This overlaps with the "more robust system", image hosting, and R7 — a shared store could solve several at once. There is no clean Microsoft-blessed API, so every path trades automation against cost/fragility.

**OUTCOME (2026-06-25): ✅ PASSED.** On Outlook for iPhone, the spike read the desktop-written `roamingSettings` value (`desktop-write-OK`) and inserted it via `setSignatureAsync` — confirming all three unknowns: (1) `roamingSettings` is readable in the iOS event runtime and syncs desktop→phone, (2) `setSignatureAsync` works on iOS, (3) `OnNewMessageCompose` fires on iPhone. **R7 (cross-platform single source) is green-lit with no backend.** The spike add-in can now be removed. Next: build the real mobile handler + the data-model changes (mobile default(s), new-vs-reply, desktop auto-insert option).
