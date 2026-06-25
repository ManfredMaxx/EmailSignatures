# Signify — Open Tasks

_Open tasks only. Remove when shipped; add a line to `03_CHANGELOG.md`._

## Now — Validate v1.3.0 + finish rollout

- [x] Enable GitHub Pages; deploy to a test user (Dan) via M365 Admin Center — done
- [ ] Dan re-tests **v1.3.0** in Outlook (reopen compose): redesigned UI, image/link URL bar, autosave, and the **send guard** (Send with *and* without a signature — should block clearly, not hang)
- [ ] End-to-end test: classic Outlook desktop (if accessible)
- [ ] Roll out to the remaining users once v1.3.0 is confirmed good

## 🛑 CRITICAL DECISION (deferred) — R0 full closure / robust send-guard

**Interim posture accepted (Dan, 2026-06-25):** keep `SendMode="PromptUser"` for now — the code already fails closed, and Dan does not want to risk destabilising his live Outlook with `Block` yet. **A more robust compliance system will ultimately be needed**; this is a deliberate, critical decision to revisit, not a quiet backlog item.

- [ ] **CRITICAL:** decide how to fully close R0 — either flip manifest to `SendMode="Block"` (+ admin `OnSendAddinsEnabled` offline policy; ≤24 h propagation), or design a more robust system. Trade-off: stability/UX vs. the residual platform fail-open hole (unsigned email can still send if the add-in is unavailable/offline). Rationale + platform facts in `04_DECISIONS.md`.
- [ ] **Constraint on the above (Dan, 2026-06-25):** whatever strict mechanism is chosen, it **must include a user-facing way to disable it** — a *prominent, off-by-default, in-pane-warning-gated* toggle. (The interim version of this toggle ships now for the PromptUser block — see "Agreed next build.")

## Now — Agreed next build (signature library + send-block toggle)

- [ ] **Multiple saved signatures** (replaces single `sigHtml` with a named list in `roamingSettings`; picker dropdown + New/Rename/**Delete with in-pane confirm**; Save/Insert/autosave act on the active one; guard total size <32 KB). Migrate existing `sigHtml` into the list.
- [ ] **Restore last saved** button (revert the open editor to its stored version).
- [ ] **Remove the "Insert starter layout" button**; seed a deletable **Example signature** in the library on first run instead.
- [ ] **Send-block disable toggle** for the current PromptUser block: prominent, **off by default**, gated by an **in-pane warning panel** (no `window.confirm`/`alert`); stored in `roamingSettings`; read by `commands.js` at send time. Consistent with R0 (conscious, warned, standing express permission).
- [ ] **Purge `window.confirm`/`alert`/`prompt`** everywhere (all silently blocked in new Outlook) — replace with in-pane UI.

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
