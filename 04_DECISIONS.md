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
