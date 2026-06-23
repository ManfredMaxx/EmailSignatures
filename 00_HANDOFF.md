# Signify — Handoff

_Last updated: 2026-06-22_

## Where we are right now

**v1.2.0 (WYSIWYG editor) is built and verified locally — not yet committed/pushed.** The form-based editor is fully replaced: `taskpane/index.html` + `taskpane/index.js` are now a freeform `contenteditable` editor with a formatting toolbar, image insert (URL + guarded paste/drag), an Insert-starter-layout button, an email-context preview, and a no-Office boot fallback. `commands.js` (send-block) is unchanged; `manifest.xml` is bumped to 1.2.0.0. README user guide rewritten for the new workflow.

Verified in a Chromium preview (matches the New Outlook/OWA engine): clean inline styling (no `<font>`/`class`), font-size→span conversion, image-by-URL, paste sanitiser, sig-marker in the insert payload, image size guard + deliverability warning, and the preview render — all pass.

**v1.1.0 (form-based) is what's currently live** at https://github.com/ManfredMaxx/EmailSignatures (public repo).

## Blocker (requires user action)

GitHub Pages has not been enabled. User must:
1. Go to github.com/ManfredMaxx/EmailSignatures → Settings → Pages
2. Source → Deploy from a branch → main / root → Save
3. Wait ~2 min for green banner

M365 deployment is blocked until Pages is live.

## What's next (in order)

1. Commit + push the v1.2.0 WYSIWYG build to GitHub
2. User enables GitHub Pages (browser-only action — see Blocker)
3. End-to-end test in real Outlook (OWA + desktop): design → Save → Insert → send-block
4. M365 Admin Center deployment to the 5 users
5. **Revisit easy image hosting** (path B) so logos "just work" (see BACKLOG / DECISIONS)

## Flagged to revisit

- **Easy no-backend image hosting** (path B). For now logos go in by URL (users host them) or as small guarded base64 pastes. The durable fix — a low-friction hosted-URL flow with no backend — is captured in `02_BACKLOG.md` and `04_DECISIONS.md`.
