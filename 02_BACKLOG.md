# Signify — Open Tasks

_Open tasks only. Remove when shipped; add a line to `03_CHANGELOG.md`._

## Now — Validate v1.3.0 + finish rollout

- [x] Enable GitHub Pages; deploy to a test user (Dan) via M365 Admin Center — done
- [ ] Dan re-tests **v1.3.0** in Outlook (reopen compose): redesigned UI, image/link URL bar, autosave, and the **send guard** (Send with *and* without a signature — should block clearly, not hang)
- [ ] End-to-end test: classic Outlook desktop (if accessible)
- [ ] Roll out to the remaining users once v1.3.0 is confirmed good

## Now — R0 full closure (Prime Directive)

- [ ] **Decision:** flip manifest `SendMode` `PromptUser` → `Block`? (fail-closed on add-in unavailable; ≤24 h propagation; needs admin offline policy). Code already fails closed; see `04_DECISIONS.md`.
- [ ] If yes: change `manifest.xml` (SendMode + version bump) and guide the admin to set `OnSendAddinsEnabled` (Exchange Online PowerShell) for offline coverage.

## Next — Mobile single source (critical requirement R7)

- [ ] **Spike:** confirm `Office.context.roamingSettings.get('sigHtml')` is readable inside the Outlook **mobile** event runtime (`OnNewMessageCompose`). If not, evaluate `sessionData` / AD-token transport-rule fallback.
- [ ] Add `<MobileFormFactor>` + `LaunchEvent OnNewMessageCompose` to `manifest.xml` (XML manifest required for mobile).
- [ ] Implement the mobile handler: read `sigHtml`, wrap + marker, `setSignatureAsync` (reuse the desktop insert path).
- [ ] Decide whether desktop/web should also auto-insert on compose (consistency) while keeping the send-block as the safety net.
- [ ] Test on Outlook iOS/Android: new + reply/forward; confirm placement and that edits on desktop reflect on mobile.

## Next — Image hosting (revisit)

- [ ] **Research easy no-backend image hosting** (plan path B) so "use my logo" resolves to a hosted URL without users needing to git-commit or find their own host. Options: an `assets/logos/` folder on the existing GitHub Pages with a low-friction upload flow (likely an admin one-time drop), or a guided in-pane flow. Rationale and trade-offs in `04_DECISIONS.md`.

## Later — Polish

- [ ] Field-tune the editor redesign and line-spacing fidelity based on Dan's feedback (complex pasted signatures with tables may still render as objects in Outlook).
