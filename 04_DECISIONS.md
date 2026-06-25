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
