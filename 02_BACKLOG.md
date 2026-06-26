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

_(Signature library, restore-last-saved, delete-with-confirm, the off-by-default send-block toggle, and the window.confirm/alert/prompt purge all **shipped in v1.4.0** — see `03_CHANGELOG.md`.)_

## Next — Mobile single source (critical requirement R7)

- [x] **Spike — ✅ PASSED (2026-06-25).** On Outlook iOS the spike read the desktop-written `roamingSettings` value and inserted via `setSignatureAsync`. Confirmed: roamingSettings readable on iOS + syncs desktop→phone, `setSignatureAsync` works, `OnNewMessageCompose` fires. **R7 green-lit, no backend.** → Remove the `spike/` add-in from M365; the `spike/` files can be deleted from the repo once the real handler lands.
- [ ] **Data model (Dan's decisions 2026-06-25):** signatures gain role assignments — desktop-new, desktop-reply, **mobile-new default, mobile-reply default** (like Outlook's native new-vs-reply, plus dedicated mobile defaults).
- [ ] **Desktop:** add an **auto-insert-on-compose option** (alongside the manual button + send-block); pick new vs reply signature by compose type.
- [ ] **Mobile handler:** on `OnNewMessageCompose`, use `getComposeTypeAsync` (mobile-supported) to pick the mobile new-vs-reply default, read from `roamingSettings`, wrap + marker, `setSignatureAsync`.
- [ ] **Manifest:** add `<MobileFormFactor>` + `OnNewMessageCompose` to production `manifest.xml` (mind the 1.12-vs-1.5 requirement-set interplay; XML manifest required).
- [ ] Test on Outlook iOS: new + reply/forward; confirm placement and that desktop edits reflect on mobile.

## Next — Image hosting (revisit)

- [ ] **Research easy no-backend image hosting** (plan path B) so "use my logo" resolves to a hosted URL without users needing to git-commit or find their own host. Options: an `assets/logos/` folder on the existing GitHub Pages with a low-friction upload flow (likely an admin one-time drop), or a guided in-pane flow. Rationale and trade-offs in `04_DECISIONS.md`.

## Landing page — shipped (2026-06-25)

✅ Live at `https://manfredmaxx.github.io/EmailSignatures/` (`index.html` at repo root). Honest positioning, editor mockup, compliance band, 4-step admin install guide; the manifest `SupportUrl` link resolves to it. Polish ideas for later: swap the CSS editor mockup for real Outlook screenshots, add an FAQ, consider a custom domain, and tune copy to Dan's taste.

## Later — Manifest exposure / distribution control (flagged by Dan, 2026-06-25)

- [ ] Revisit how locked-down the manifest should be. The link was removed from the **landing page** (done), but it still appears in the public **README** and the manifest stays fetchable at its URL (public repo). Options when revisited: (a) also strip from README; (b) rethink hosting for true privacy (a private repo breaks free GitHub Pages → would mean moving hosting). Note: a leaked manifest is a distribution/branding concern, **not** data exposure — each user's signatures stay in their own mailbox. Reason for the urgency wasn't given; confirm with Dan when picked up.

## Later — Polish

- [ ] Field-tune the editor redesign and line-spacing fidelity based on Dan's feedback (complex pasted signatures with tables may still render as objects in Outlook).
