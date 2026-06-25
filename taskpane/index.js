'use strict';

// ════════════════════════════════════════════════════════════════════════════
//  Signify — WYSIWYG signature editor with a signature library.
//
//  Storage (per-user, Exchange mailbox via roamingSettings):
//    sigLibrary        = { v, activeId, items:[{id,name,html}] }
//    sendBlockDisabled = bool  (the off-by-default send-guard opt-out)
//  A per-signature working copy is auto-kept in localStorage so nothing is
//  lost if Save is forgotten. No window.confirm/alert/prompt anywhere — all are
//  silently blocked in new Outlook; we use in-pane UI instead.
// ════════════════════════════════════════════════════════════════════════════

// ── Constants ────────────────────────────────────────────────────────────────

var LIB_KEY      = 'sigLibrary';
var BLOCK_KEY    = 'sendBlockDisabled';
var STORAGE_KEY  = 'sigHtml';            // legacy single signature (migrated then dropped)
var LEGACY_KEY   = 'sigFields';          // legacy form object (migrated then dropped)
var DRAFT_PREFIX = 'signify_draft_';     // + signature id  → localStorage working copy

var MARKER = '<div id="sig-marker" style="display:none;height:0;line-height:0;overflow:hidden;">&#8203;</div>';

var MAX_IMG_CHARS = 20000;               // single embedded image ceiling
var MAX_LIB_CHARS = 30000;               // total library JSON, headroom under ~32 KB cap

var SWATCHES = ['#000000','#444444','#777777','#0f6cbd','#106ebe','#005a9e',
                '#b3261e','#1a7f37','#7a3ea1','#b07b00','#c2185b','#00188f'];

var STARTER_TEMPLATE =
  '<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#333333;line-height:1.4;">' +
    '<div style="font-size:15px;font-weight:bold;color:#000000;">Your Name</div>' +
    '<div style="font-size:12px;color:#555555;padding-top:2px;">Job Title &nbsp;&bull;&nbsp; Company Name</div>' +
    '<div style="border-top:1px solid #dddddd;margin:7px 0;font-size:1px;line-height:1px;">&nbsp;</div>' +
    '<div style="font-size:12px;color:#333333;">' +
      'Office: 000-000-0000 &nbsp;&nbsp;|&nbsp;&nbsp; ' +
      '<a href="mailto:you@company.com" style="color:#0f6cbd;text-decoration:none;">you@company.com</a> &nbsp;&nbsp;|&nbsp;&nbsp; ' +
      '<a href="https://company.com" style="color:#0f6cbd;text-decoration:none;">company.com</a>' +
    '</div>' +
    '<div style="margin-top:8px;padding:5px 8px;background-color:#f5f5f5;border-left:2px solid #0f6cbd;font-size:11px;color:#444444;">' +
      'Licensed Professional #00000 &nbsp;&bull;&nbsp; Regulated by [Agency]. Replace with your required compliance text.' +
    '</div>' +
    '<div style="margin-top:6px;font-size:9px;color:#aaaaaa;line-height:1.4;">' +
      'This email and any attachments are confidential and intended solely for the addressee. ' +
      'If you received this in error, please delete it and notify the sender.' +
    '</div>' +
  '</div>';

// ── State ────────────────────────────────────────────────────────────────────

var library = { v: 1, activeId: null, items: [] };
var sendBlockDisabled = false;

var editor = null, booted = false, devMode = false, savedRange = null,
    noteShown = false, urlMode = null, confirmCallback = null;
var _previewTimer = null, _statusTimer = null, _autosaveTimer = null;

// ── Startup ──────────────────────────────────────────────────────────────────

if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady(function (info) { boot(!(info && info.host === Office.HostType.Outlook)); });
}
setTimeout(function () { if (!booted) boot(true); }, 1500);

function boot(isDev) {
  if (booted) return;
  booted = true; devMode = isDev;
  wireEditor();
  buildSwatches();
  loadLibrary();
  renderPicker();
  loadActiveIntoEditor();
  refreshProtectionUI();
  updatePreview();
  document.getElementById('loading').style.display = 'none';
  document.getElementById('app').style.display = 'flex';
  if (devMode) document.getElementById('devNotice').style.display = 'block';
}

function hasRoaming() {
  try { return !devMode && !!(Office.context && Office.context.roamingSettings); }
  catch (e) { return false; }
}

function newId() {
  return 'sig_' + Date.now().toString(36) + '_' + Math.floor(Math.random() * 1e6).toString(36);
}

// ── Library load / migrate / persist ─────────────────────────────────────────

function loadLibrary() {
  if (hasRoaming()) {
    try { sendBlockDisabled = Office.context.roamingSettings.get(BLOCK_KEY) === true; } catch (e) {}
  }
  var lib = null;
  if (hasRoaming()) {
    try {
      var raw = Office.context.roamingSettings.get(LIB_KEY);
      if (raw) lib = typeof raw === 'string' ? JSON.parse(raw) : raw;
    } catch (e) {}
  }
  if (lib && lib.items && lib.items.length) { library = normalizeLib(lib); return; }

  var migrated = migrateLegacy();
  var id = newId();
  library = migrated
    ? { v: 1, activeId: id, items: [{ id: id, name: 'My signature', html: migrated }] }
    : { v: 1, activeId: id, items: [{ id: id, name: 'Example signature', html: STARTER_TEMPLATE }] };
}

function normalizeLib(lib) {
  lib.v = 1;
  lib.items = (lib.items || []).filter(function (it) { return it && it.id; });
  if (!lib.items.length) { var id = newId(); lib.items = [{ id: id, name: 'Signature', html: '' }]; lib.activeId = id; }
  if (!lib.items.some(function (it) { return it.id === lib.activeId; })) lib.activeId = lib.items[0].id;
  return lib;
}

function migrateLegacy() {
  if (!hasRoaming()) return null;
  try {
    var html = Office.context.roamingSettings.get(STORAGE_KEY);
    if (typeof html === 'string' && html.trim()) return html;
    var legacy = Office.context.roamingSettings.get(LEGACY_KEY);
    if (legacy) { var f = typeof legacy === 'string' ? JSON.parse(legacy) : legacy; return legacyHtml(f); }
  } catch (e) {}
  return null;
}

function persistLibrary(onDone) {
  if (!hasRoaming()) { if (onDone) onDone(true); return; }
  var json = JSON.stringify(library);
  if (json.length > MAX_LIB_CHARS) {
    showStatus('Your signatures are too large to save together — likely an embedded image. ' +
      'Add logos with the image button (by web address) instead.', 'error');
    if (onDone) onDone(false); return;
  }
  try {
    Office.context.roamingSettings.set(LIB_KEY, json);
    Office.context.roamingSettings.remove(STORAGE_KEY);
    Office.context.roamingSettings.remove(LEGACY_KEY);
    Office.context.roamingSettings.saveAsync(function (r) {
      var ok = r.status === Office.AsyncResultStatus.Succeeded;
      if (!ok) {
        var msg = (r.error && r.error.message) || '';
        if (/9057/.test(msg) || /size/i.test(msg)) showStatus('Too large to save — use logos by web address.', 'error');
        else showStatus('Save failed — check your connection. (Your work is kept on this device.)', 'error');
      }
      if (onDone) onDone(ok);
    });
  } catch (e) { showStatus('Save failed: ' + e.message, 'error'); if (onDone) onDone(false); }
}

function activeItem() {
  for (var i = 0; i < library.items.length; i++) if (library.items[i].id === library.activeId) return library.items[i];
  return null;
}

// ── Picker + switching ───────────────────────────────────────────────────────

function renderPicker() {
  var sel = document.getElementById('sigPicker');
  sel.innerHTML = '';
  library.items.forEach(function (it) {
    var o = document.createElement('option');
    o.value = it.id; o.textContent = it.name || 'Untitled';
    if (it.id === library.activeId) o.selected = true;
    sel.appendChild(o);
  });
}

function selectSignature(id) {
  var cur = activeItem();
  if (cur) setDraft(cur.id, editor.innerHTML);   // keep current edits
  library.activeId = id;
  loadActiveIntoEditor();
  renderPicker();
  closeSubbars();
}

function loadActiveIntoEditor() {
  var it = activeItem();
  if (!it) { editor.innerHTML = ''; setSaveState('clean'); return; }
  var draft = getDraft(it.id);
  if (draft != null && draft !== it.html) { editor.innerHTML = draft; setSaveState('unsaved'); }
  else { editor.innerHTML = it.html; setSaveState('clean'); }
  schedulePreview();
}

// ── New / Rename / Delete ────────────────────────────────────────────────────

function newSignature() {
  var cur = activeItem();
  if (cur) setDraft(cur.id, editor.innerHTML);
  var id = newId();
  library.items.push({ id: id, name: uniqueName('Signature'), html: STARTER_TEMPLATE });
  library.activeId = id;
  persistLibrary(function () {
    loadActiveIntoEditor(); renderPicker();
    showStatus('New signature added. Rename it with the ✎ button, edit, then Save.', 'info');
  });
}

function uniqueName(base) {
  var n = library.items.length + 1;
  var names = library.items.map(function (it) { return it.name; });
  var name = base + ' ' + n;
  while (names.indexOf(name) !== -1) { n++; name = base + ' ' + n; }
  return name;
}

function openRename() {
  var it = activeItem(); if (!it) return;
  closeSubbars();
  var inp = document.getElementById('renameInput');
  inp.value = it.name || '';
  document.getElementById('renameBar').classList.add('show');
  setTimeout(function () { inp.focus(); inp.select(); }, 0);
}
function renameKey(e) {
  if (e.key === 'Enter') { e.preventDefault(); confirmRename(); }
  else if (e.key === 'Escape') { e.preventDefault(); closeSubbars(); }
}
function confirmRename(e) {
  if (e) e.preventDefault();
  var it = activeItem(); if (!it) return;
  var name = (document.getElementById('renameInput').value || '').trim();
  if (!name) { showStatus('Enter a name.', 'info'); return; }
  it.name = name.slice(0, 40);
  persistLibrary(function () { renderPicker(); });
  closeSubbars();
  showStatus('Renamed.', 'success');
}

function askDelete() {
  var it = activeItem(); if (!it) return;
  showConfirm('Delete signature', 'Delete “' + (it.name || 'Untitled') + '”? This can’t be undone.',
    'Delete', true, function () { doDelete(it.id); });
}
function doDelete(id) {
  clearDraft(id);
  library.items = library.items.filter(function (it) { return it.id !== id; });
  if (!library.items.length) {
    var nid = newId();
    library.items.push({ id: nid, name: 'Example signature', html: STARTER_TEMPLATE });
    library.activeId = nid;
  } else if (library.activeId === id) {
    library.activeId = library.items[0].id;
  }
  persistLibrary(function () {
    loadActiveIntoEditor(); renderPicker();
    showStatus('Signature deleted.', 'success');
  });
}

// ── Save / Insert / Restore ──────────────────────────────────────────────────

function saveSignature() {
  if (!hasRoaming()) { showStatus('Save works inside Outlook. (Preview mode.)', 'warn'); return; }
  var it = activeItem();
  if (!it) { showStatus('No signature selected.', 'error'); return; }
  it.html = editor.innerHTML.trim();
  library.activeId = it.id;
  persistLibrary(function (ok) {
    if (ok) { clearDraft(it.id); setSaveState('saved'); showStatus('Saved.', 'success'); }
  });
}

function insertSignature() {
  if (devMode) { showStatus('Insert works inside Outlook. (Preview mode.)', 'warn'); return; }
  var inner = editor.innerHTML.trim();
  if (!inner) { showStatus('This signature is empty — add something first.', 'error'); return; }
  var item = (Office.context.mailbox && Office.context.mailbox.item) || null;
  if (!item) { showStatus('No compose window found. Open a new email first.', 'error'); return; }
  var html = wrapSignature(inner) + MARKER;
  var opts = { coercionType: Office.CoercionType.Html };
  if (typeof item.body.setSignatureAsync === 'function') item.body.setSignatureAsync(html, opts, function (r) { insertDone(r, false); });
  else item.body.prependAsync(html, opts, function (r) { insertDone(r, true); });
}
function insertDone(r, compat) {
  if (r.status === Office.AsyncResultStatus.Succeeded) showStatus(compat ? 'Signature inserted (compatibility mode).' : 'Signature inserted.', 'success');
  else showStatus('Insert failed: ' + ((r.error && r.error.message) || 'unknown error'), 'error');
}

function restoreLastSaved() {
  var it = activeItem(); if (!it) return;
  editor.innerHTML = it.html;
  clearDraft(it.id);
  setSaveState('clean');
  schedulePreview();
  showStatus('Restored the last saved version.', 'info');
}

function wrapSignature(inner) {
  return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#333333;line-height:1.4;">' + inner + '</div>';
}

// ── Drafts (per-signature local working copy) ────────────────────────────────

function draftKey(id) { return DRAFT_PREFIX + id; }
function getDraft(id) { try { return localStorage.getItem(draftKey(id)); } catch (e) { return null; } }
function setDraft(id, html) { try { localStorage.setItem(draftKey(id), html); } catch (e) {} }
function clearDraft(id) { try { localStorage.removeItem(draftKey(id)); } catch (e) {} }

function markDirty() {
  var it = activeItem(); if (!it) return;
  setSaveState('unsaved');
  clearTimeout(_autosaveTimer);
  _autosaveTimer = setTimeout(function () { setDraft(it.id, editor.innerHTML); }, 500);
}

function setSaveState(state) {
  var t = document.getElementById('saveStateText');
  var r = document.getElementById('btnRestore');
  if (!t) return;
  t.className = 'state ' + (state === 'unsaved' ? 'unsaved' : (state === 'saved' ? 'saved' : ''));
  if (state === 'unsaved') { t.textContent = '● Unsaved changes (kept on this device)'; r.style.display = 'inline'; }
  else if (state === 'saved') { t.textContent = '✓ Saved to your account'; r.style.display = 'none'; }
  else { t.textContent = ''; r.style.display = 'none'; }
}

// ── Send-guard protection toggle ─────────────────────────────────────────────

function refreshProtectionUI() {
  var btn = document.getElementById('protectBtn');
  var sub = document.getElementById('protectSub');
  var banner = document.getElementById('blockOffBanner');
  if (sendBlockDisabled) {
    btn.classList.add('off'); btn.setAttribute('aria-checked', 'false');
    sub.textContent = 'OFF — emails can send without a signature';
    banner.style.display = 'flex';
  } else {
    btn.classList.remove('off'); btn.setAttribute('aria-checked', 'true');
    sub.textContent = 'On — emails are blocked without a signature';
    banner.style.display = 'none';
  }
}
function toggleProtection() {
  if (!sendBlockDisabled) {
    showConfirm('Turn off the signature check?',
      'Emails will be able to send without a signature. Only do this if you understand the risk — ' +
      'your signatures (and any required compliance text) won’t be enforced.',
      'Turn it off', true, function () { setProtection(true); });
  } else {
    setProtection(false);
  }
}
function protectionOn() { setProtection(false); }
function setProtection(disabled) {
  sendBlockDisabled = disabled;
  refreshProtectionUI();
  if (hasRoaming()) {
    try { Office.context.roamingSettings.set(BLOCK_KEY, disabled); Office.context.roamingSettings.saveAsync(function () {}); } catch (e) {}
  }
  showStatus(disabled ? 'Signature check turned OFF — emails can send without a signature.' : 'Signature check turned back on.',
    disabled ? 'warn' : 'success');
}

// ── In-pane confirm (replaces window.confirm) ────────────────────────────────

function showConfirm(title, msg, okLabel, danger, onOk) {
  document.getElementById('confirmTitle').textContent = title;
  document.getElementById('confirmMsg').textContent = msg;
  var ok = document.getElementById('confirmOk');
  ok.textContent = okLabel || 'Confirm';
  ok.className = 'ok' + (danger ? ' danger' : '');
  confirmCallback = onOk || null;
  ok.onclick = function () { var cb = confirmCallback; closeConfirm(); if (cb) cb(); };
  document.getElementById('confirmOverlay').classList.add('show');
}
function closeConfirm() {
  document.getElementById('confirmOverlay').classList.remove('show');
  confirmCallback = null;
}

// ── Editor wiring ────────────────────────────────────────────────────────────

function wireEditor() {
  editor = document.getElementById('editor');
  editor.addEventListener('keyup', function () { saveRange(); });
  editor.addEventListener('mouseup', saveRange);
  editor.addEventListener('input', function () { markDirty(); schedulePreview(); });
  editor.addEventListener('paste', onPaste);
  editor.addEventListener('dragover', function (e) { e.preventDefault(); editor.classList.add('drag-over'); });
  editor.addEventListener('dragleave', function () { editor.classList.remove('drag-over'); });
  editor.addEventListener('drop', onDrop);
  try {
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('defaultParagraphSeparator', false, 'div');
  } catch (e) {}
}

// ── Selection tracking ───────────────────────────────────────────────────────

function saveRange() {
  var sel = window.getSelection();
  if (sel && sel.rangeCount && editor.contains(sel.anchorNode)) savedRange = sel.getRangeAt(0).cloneRange();
}
function restoreRange() {
  if (!savedRange) return false;
  try { var s = window.getSelection(); s.removeAllRanges(); s.addRange(savedRange); return true; } catch (e) { return false; }
}
function caretToEnd() {
  try { var r = document.createRange(); r.selectNodeContents(editor); r.collapse(false); var s = window.getSelection(); s.removeAllRanges(); s.addRange(r); } catch (e) {}
}
function focusEditor() { editor.focus(); if (!restoreRange()) caretToEnd(); }

// ── Toolbar ──────────────────────────────────────────────────────────────────

function cmd(e, command) {
  if (e) e.preventDefault();
  editor.focus();
  try { document.execCommand('styleWithCSS', false, true); document.execCommand(command, false, null); } catch (err) {}
  saveRange(); markDirty(); schedulePreview();
}
function applyFont(family) {
  if (!family) return;
  focusEditor();
  try { document.execCommand('styleWithCSS', false, true); document.execCommand('fontName', false, family); } catch (e) {}
  document.getElementById('fontSelect').value = '';
  saveRange(); markDirty(); schedulePreview();
}
function applySize(px) {
  if (!px) return;
  focusEditor();
  var sel = window.getSelection();
  if (!sel || sel.isCollapsed) { showStatus('Select some text first, then choose a size.', 'info'); document.getElementById('sizeSelect').value = ''; return; }
  try { document.execCommand('styleWithCSS', false, false); document.execCommand('fontSize', false, '7'); document.execCommand('styleWithCSS', false, true); } catch (e) {}
  toArray(editor.querySelectorAll('font[size="7"]')).forEach(function (f) {
    var span = document.createElement('span'); span.style.fontSize = px + 'px';
    while (f.firstChild) span.appendChild(f.firstChild);
    f.parentNode.replaceChild(span, f);
  });
  toArray(editor.querySelectorAll('span[style*="xxx-large"]')).forEach(function (s) { s.style.fontSize = px + 'px'; });
  document.getElementById('sizeSelect').value = '';
  saveRange(); markDirty(); schedulePreview();
}

function buildSwatches() {
  var wrap = document.getElementById('swatches');
  if (!wrap) return;
  SWATCHES.forEach(function (c) {
    var b = document.createElement('button');
    b.type = 'button'; b.className = 'swatch'; b.style.background = c; b.title = c;
    b.addEventListener('mousedown', function (ev) { ev.preventDefault(); pickColor(c); });
    wrap.appendChild(b);
  });
  var custom = document.createElement('label');
  custom.className = 'swatch custom'; custom.title = 'Custom colour';
  var inp = document.createElement('input'); inp.type = 'color'; inp.value = '#333333';
  inp.addEventListener('input', function () { pickColor(inp.value); });
  custom.appendChild(inp); wrap.appendChild(custom);
}
function toggleColors(e) {
  if (e) e.preventDefault();
  saveRange();
  document.getElementById('urlBar').classList.remove('show');
  document.getElementById('renameBar').classList.remove('show');
  document.getElementById('colorBar2').classList.toggle('show');
}
function pickColor(color) {
  focusEditor();
  try { document.execCommand('styleWithCSS', false, true); document.execCommand('foreColor', false, color); } catch (e) {}
  var bar = document.getElementById('colorBar'); if (bar) bar.style.background = color;
  saveRange(); markDirty(); schedulePreview(); closeSubbars();
}

// ── Image / Link via inline URL bar ──────────────────────────────────────────

function openImage(e) { if (e) e.preventDefault(); openUrlBar('image'); }
function openLink(e) { if (e) e.preventDefault(); openUrlBar('link'); }
function openUrlBar(mode) {
  saveRange(); urlMode = mode;
  document.getElementById('colorBar2').classList.remove('show');
  document.getElementById('renameBar').classList.remove('show');
  document.getElementById('urlLabel').textContent = mode === 'image' ? 'Image web address' : 'Link web address';
  var inp = document.getElementById('urlInput');
  inp.value = '';
  inp.placeholder = mode === 'image' ? 'https://…  (a link to a logo or image file)' : 'https://…';
  document.getElementById('urlBar').classList.add('show');
  setTimeout(function () { inp.focus(); }, 0);
}
function urlKey(e) {
  if (e.key === 'Enter') { e.preventDefault(); confirmUrl(); }
  else if (e.key === 'Escape') { e.preventDefault(); closeSubbars(); }
}
function confirmUrl(e) {
  if (e) e.preventDefault();
  var url = (document.getElementById('urlInput').value || '').trim();
  if (!url) { showStatus('Enter a web address first.', 'info'); return; }
  if (urlMode === 'image') {
    if (!/^https:\/\//i.test(url)) { showStatus('Use a secure image address — it must start with https://', 'error'); return; }
    insertHtmlAtCursor('<img src="' + esc(url) + '" style="max-width:100%;height:auto;" alt=""/>');
  } else {
    if (!/^https?:\/\//i.test(url) && !/^mailto:/i.test(url)) url = 'https://' + url;
    focusEditor();
    var sel = window.getSelection();
    if (!sel || sel.isCollapsed) insertHtmlAtCursor('<a href="' + esc(url) + '" style="color:#0f6cbd;">' + esc(url) + '</a>');
    else { try { document.execCommand('styleWithCSS', false, true); document.execCommand('createLink', false, url); } catch (err) {} saveRange(); markDirty(); schedulePreview(); }
  }
  closeSubbars();
}
function closeSubbars(e) {
  if (e) e.preventDefault();
  document.getElementById('urlBar').classList.remove('show');
  document.getElementById('colorBar2').classList.remove('show');
  document.getElementById('renameBar').classList.remove('show');
  urlMode = null;
}

// ── Image paste / drop ───────────────────────────────────────────────────────

function onPaste(e) {
  var dt = e.clipboardData; if (!dt) return;
  var file = imageFromItems(dt.items);
  if (file) { e.preventDefault(); handleImageFile(file); return; }
  var html = dt.getData && dt.getData('text/html');
  if (html) { e.preventDefault(); insertHtmlAtCursor(sanitizeHtml(html)); }
}
function onDrop(e) {
  editor.classList.remove('drag-over');
  var files = e.dataTransfer && e.dataTransfer.files;
  if (!files || !files.length) return;
  for (var i = 0; i < files.length; i++) {
    if (files[i].type && files[i].type.indexOf('image/') === 0) { e.preventDefault(); handleImageFile(files[i]); return; }
  }
}
function imageFromItems(items) {
  if (!items) return null;
  for (var i = 0; i < items.length; i++) if (items[i].kind === 'file' && items[i].type.indexOf('image/') === 0) return items[i].getAsFile();
  return null;
}
function handleImageFile(file) {
  var reader = new FileReader();
  reader.onload = function () {
    var dataUrl = String(reader.result || '');
    if (dataUrl.length > MAX_IMG_CHARS) {
      showStatus('That image is too large to save with your signature (about ' + Math.round(dataUrl.length / 1024) +
        ' KB; limit ~20 KB). Host it online and add it with the image button (by web address) instead.', 'error');
      return;
    }
    insertHtmlAtCursor('<img src="' + dataUrl + '" style="max-width:100%;height:auto;" alt=""/>');
    deliverabilityNote();
  };
  reader.readAsDataURL(file);
}
function deliverabilityNote() {
  if (noteShown) return; noteShown = true;
  showStatus('Image added. Heads-up: pasted images may not display in Gmail and some clients — ' +
    'for a logo, the image button (by web address) is the most reliable.', 'warn');
}
function insertHtmlAtCursor(html) {
  focusEditor();
  try { document.execCommand('insertHTML', false, html); } catch (e) { editor.innerHTML += html; }
  saveRange(); markDirty(); schedulePreview();
}

// ── Paste sanitizer ──────────────────────────────────────────────────────────

function sanitizeHtml(html) {
  var tmp = document.createElement('div'); tmp.innerHTML = html;
  toArray(tmp.querySelectorAll('script, style, meta, link, title')).forEach(function (n) { n.parentNode && n.parentNode.removeChild(n); });
  toArray(tmp.getElementsByTagName('*')).forEach(function (el) {
    if (el.tagName && el.tagName.indexOf(':') !== -1) { while (el.firstChild) el.parentNode.insertBefore(el.firstChild, el); el.parentNode.removeChild(el); }
  });
  toArray(tmp.querySelectorAll('*')).forEach(function (el) {
    toArray(el.attributes).forEach(function (attr) {
      var name = attr.name.toLowerCase();
      if (name === 'class' || name === 'id' || name.indexOf('on') === 0 || name === 'lang' || name === 'align' ||
          name.indexOf('xmlns') === 0 || name.indexOf('mso') !== -1 || name.indexOf('data-') === 0) el.removeAttribute(attr.name);
    });
    var style = el.getAttribute && el.getAttribute('style');
    if (style) { var cleaned = style.replace(/mso-[^;]+;?/gi, '').replace(/^\s+|\s+$/g, ''); if (cleaned) el.setAttribute('style', cleaned); else el.removeAttribute('style'); }
  });
  return tmp.innerHTML;
}

// ── Tabs + preview ───────────────────────────────────────────────────────────

function showTab(name) {
  var tabs = document.querySelectorAll('.tab-btn');
  for (var i = 0; i < tabs.length; i++) tabs[i].classList.toggle('active', tabs[i].dataset.tab === name);
  document.getElementById('tab-edit').style.display = name === 'edit' ? 'flex' : 'none';
  document.getElementById('tab-preview').style.display = name === 'preview' ? 'flex' : 'none';
  if (name === 'preview') updatePreview();
}
function schedulePreview() { clearTimeout(_previewTimer); _previewTimer = setTimeout(updatePreview, 150); }
function updatePreview() {
  var el = document.getElementById('previewHtml'); if (!el) return;
  var inner = editor ? editor.innerHTML.trim() : '';
  el.innerHTML = inner ? wrapSignature(inner) : '<p style="color:#9aa3af;font-size:13px;margin:0;">Your signature preview will appear here.</p>';
}

// ── Legacy migration renderer ────────────────────────────────────────────────

function legacyHtml(f) {
  if (!f) return '';
  var accent = f.accentColor || '#0f6cbd';
  var p = [];
  if (f.fullName) p.push('<div style="font-size:15px;font-weight:bold;color:#000;">' + esc(f.fullName) + '</div>');
  var meta = [f.jobTitle, f.company, f.department].filter(Boolean).map(esc);
  if (meta.length) p.push('<div style="font-size:12px;color:#555;padding-top:2px;">' + meta.join(' &bull; ') + '</div>');
  var c = [];
  if (f.phone) c.push('Office: ' + esc(f.phone));
  if (f.mobile) c.push('Mobile: ' + esc(f.mobile));
  if (f.email) c.push('<a href="mailto:' + esc(f.email) + '" style="color:' + accent + ';text-decoration:none;">' + esc(f.email) + '</a>');
  if (f.website) c.push('<a href="https://' + esc(f.website) + '" style="color:' + accent + ';text-decoration:none;">' + esc(f.website) + '</a>');
  if (c.length) p.push('<div style="font-size:12px;color:#333;padding-top:6px;">' + c.join(' &nbsp;|&nbsp; ') + '</div>');
  if (f.address) p.push('<div style="font-size:11px;color:#777;padding-top:4px;">' + esc(f.address).replace(/\n/g, '<br>') + '</div>');
  if (f.compliance) p.push('<div style="margin-top:8px;padding:5px 8px;background:#f5f5f5;border-left:2px solid ' + accent + ';font-size:11px;color:#444;">' + esc(f.compliance).replace(/\n/g, '<br>') + '</div>');
  if (f.disclaimer) p.push('<div style="margin-top:6px;font-size:9px;color:#aaa;">' + esc(f.disclaimer).replace(/\n/g, '<br>') + '</div>');
  return '<div style="font-family:Arial,Helvetica,sans-serif;">' + p.join('') + '</div>';
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function toArray(nodeList) { return Array.prototype.slice.call(nodeList || []); }
function esc(str) { if (!str) return ''; return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function showStatus(msg, type) {
  var el = document.getElementById('statusMsg'); if (!el) return;
  el.textContent = msg; el.className = 'status-msg status-' + (type || 'info'); el.style.display = 'block';
  clearTimeout(_statusTimer);
  _statusTimer = setTimeout(function () { el.style.display = 'none'; }, (type === 'warn' || type === 'error') ? 7000 : 4000);
}
