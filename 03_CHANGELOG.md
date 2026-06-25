# Signify — Changelog

## v1.4.0 — 2026-06-25 (signature library + send-block toggle; code-only, manifest unchanged)
- **Multiple saved signatures.** Storage moved from a single `sigHtml` to a named library (`sigLibrary` in roamingSettings): a picker dropdown plus **New / Rename / Delete** (delete uses an in-pane confirm). Save / Insert / autosave all act on the selected signature. Existing `sigHtml` / `sigFields` migrate into the library automatically.
- **Restore last saved** — revert the open editor to its stored version (discards the local draft).
- **Per-signature auto-save** — each signature's working copy is kept in `localStorage` and restored on reopen; switching signatures preserves unsaved edits.
- **Starter button removed** — a deletable "Example signature" is seeded on first run instead (and New signatures start from that template), so the library demonstrates itself.
- **Send-block disable toggle** — a prominent, **off-by-default** "Require a signature to send" switch, gated by an in-pane warning; stored in roamingSettings (`sendBlockDisabled`) and honoured by the send guard (`commands.js`). A loud banner shows whenever protection is off. Consistent with R0 (conscious, warned, standing express permission).
- **Purged `window.confirm` / `alert` / `prompt`** everywhere — all silently blocked in new Outlook (this is what made the old starter button and URL dialog do nothing); replaced with in-pane UI.
- The library's total size is guarded (<32 KB); embedded images still count against it — durable fix is hosted-URL images (image-hosting build issue, R6, in `04_DECISIONS.md`).

## v1.3.0 — 2026-06-25 (editor redesign + R0 send-guard hardening; code-only, manifest unchanged)
- **R0 / Prime Directive established** (see `REQUIREMENTS.md`). The send guard now **fails CLOSED**: `commands.js` no longer allows a send when it can't verify a signature (was fail-open on read error — a silent-unsigned-send hole). Added `Office.onReady` init (fixes the hang), a ~4 s completion safety net that **blocks** rather than hanging, and a runtime "Send Anyway" override offered **only** when a signature is genuinely missing. Full closure (manifest `SendMode="Block"` + admin offline policy) is **pending Dan's decision** — manifest intentionally unchanged in this release.
- **Editor UI redesigned** — roomier, modern toolbar (36 px controls, clearer grouping, refined palette and spacing); larger editor and footer.
- **Image / Link by URL fixed** — replaced `window.prompt` (silently blocked in new Outlook) with an **inline URL input bar**.
- **Auto-save** — the working copy is kept in `localStorage` and restored on reopen, with an unsaved/saved indicator; work is no longer lost if the user forgets to Save.
- **Colour picker** — preset swatch palette + custom, replacing the raw OS colour dialog.
- **Line spacing** — `defaultParagraphSeparator=div` + zeroed editor block margins for tighter, email-faithful spacing; starter layout rebuilt with `<div>`s. (Pasted complex/3rd-party signatures with tables may still render as objects in Outlook — building in the editor is most faithful.)

## v1.2.0 — 2026-06-22
- **Replaced the form-based editor with a full WYSIWYG rich-text editor.** Users now design their signature freely (no fixed template) in a `contenteditable` canvas with a formatting toolbar: bold/italic/underline, font family, font size, text colour, alignment, bullet list, link, and clear-formatting.
- Font size and colour emit clean inline `<span style>` (no `<font>` tags) via `styleWithCSS` + a size-7→px conversion, so output survives email clients.
- **Images:** insert by URL (recommended for logos) plus paste and drag-and-drop. Pasted/dragged images embed as base64 with a ~20 KB size guard and a one-time deliverability warning (base64 doesn't render in Gmail; large images can't be saved).
- **Insert starter layout** button: one-click sample signature (name/title/contact/compliance box/disclaimer) the user can edit or delete.
- Light paste sanitiser strips `<script>`/`<style>`, `class`/`id`/`on*`/`mso-` junk and Word namespaced tags while keeping inline styles.
- Storage moved from the `sigFields` object to a raw-HTML `sigHtml` key (with one-time migration of any old form data). Save guards total size (~30 KB) and handles roamingSettings error 9057.
- Insert still appends the hidden `sig-marker` (send-block contract unchanged); `commands.js` untouched.
- Added a no-Office boot fallback so the editor renders and is testable in a plain browser. Manifest version → 1.2.0.0.
- **Made the manifest pass M365 Admin Center validation.** Added base `<IconUrl>` (64px) + `<HighResolutionIconUrl>` (128px) with generated `assets/icon-64.png`/`icon-128.png`; removed the schema-invalid group-level `<Icon>` and the `<TaskpaneId>` from the `ShowTaskpane` action; corrected `bt:Image` `Size`→`size` casing; set a real `ProviderName`. Verified clean with `office-addin-manifest validate`.

## v1.1.0 — 2026-06-22
- Rebranded all user-facing text to "Signify" (manifest DisplayName, ribbon button, task pane header, send-block error message)
- Added comprehensive `README.md`: admin setup guide (GitHub Pages + M365 Admin Center) and complete user guide
- Made GitHub repo public to enable free GitHub Pages hosting

## v1.0.0 — 2026-06-22
- Initial build: form-based signature editor in Outlook task pane
- Send-block via `OnMessageSend` with `SendMode="PromptUser"`
- Signature storage via `roamingSettings` (no backend required)
- `setSignatureAsync` for correct placement (after new text, before quoted thread)
- Hidden `<div id="sig-marker">` as send-block detection mechanism
- GitHub Pages hosting: five static files, no server
- Icons generated via browser-based canvas tool (`assets/generate-icons.html`)
