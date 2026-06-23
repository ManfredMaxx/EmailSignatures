# Signify — Changelog

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
