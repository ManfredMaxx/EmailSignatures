# Signify — Open Tasks

_Open tasks only. Remove when shipped; add a line to `03_CHANGELOG.md`._

## Now — Deployment

- [ ] Enable GitHub Pages (user/admin browser step) and verify the URL returns 200 for `manifest.xml`
- [ ] End-to-end test in real Outlook: OWA compose → open Signify → design → Save → Insert → send-block check
- [ ] End-to-end test: classic Outlook desktop (if accessible)
- [ ] Deploy to 5 users via M365 Admin Center

## Next — Mobile single source (critical requirement R7)

- [ ] **Spike:** confirm `Office.context.roamingSettings.get('sigHtml')` is readable inside the Outlook **mobile** event runtime (`OnNewMessageCompose`). If not, evaluate `sessionData` / AD-token transport-rule fallback.
- [ ] Add `<MobileFormFactor>` + `LaunchEvent OnNewMessageCompose` to `manifest.xml` (XML manifest required for mobile).
- [ ] Implement the mobile handler: read `sigHtml`, wrap + marker, `setSignatureAsync` (reuse the desktop insert path).
- [ ] Decide whether desktop/web should also auto-insert on compose (consistency) while keeping the send-block as the safety net.
- [ ] Test on Outlook iOS/Android: new + reply/forward; confirm placement and that edits on desktop reflect on mobile.

## Next — Image hosting (revisit)

- [ ] **Research easy no-backend image hosting** (plan path B) so "use my logo" resolves to a hosted URL without users needing to git-commit or find their own host. Options: an `assets/logos/` folder on the existing GitHub Pages with a low-friction upload flow (likely an admin one-time drop), or a guided in-pane flow. Rationale and trade-offs in `04_DECISIONS.md`.

## Later — Polish

- [ ] Consider replacing `window.prompt` for image/link URLs with an inline in-pane input field — `prompt` can be unreliable in some Outlook task-pane hosts.
