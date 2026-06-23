'use strict';

// ════════════════════════════════════════════════════════════════════════════
//  Signify — WYSIWYG signature editor
//  The contenteditable #editor is the live signature. Toolbar actions use
//  document.execCommand (with styleWithCSS so output is inline-styled, not
//  <font> tags). The raw HTML is stored per-user in roamingSettings under
//  'sigHtml'. Insert wraps it in a base-font div and appends the hidden
//  sig-marker that the send-block (commands.js) looks for.
// ════════════════════════════════════════════════════════════════════════════

// ── Constants ────────────────────────────────────────────────────────────────

var STORAGE_KEY = 'sigHtml';        // new: raw signature HTML
var LEGACY_KEY  = 'sigFields';      // old: form-field object (migrated then dropped)

// Hidden marker appended at insert time so the user can never delete it.
// The zero-width space keeps the element from being trimmed by Outlook.
var MARKER = '<div id="sig-marker" style="display:none;height:0;line-height:0;overflow:hidden;">&#8203;</div>';

var MAX_IMG_CHARS = 20000;          // ~14 KB raw — a single embedded image ceiling
var MAX_SIG_CHARS = 30000;          // headroom under the ~32 KB roamingSettings cap

// ── State ────────────────────────────────────────────────────────────────────

var editor      = null;
var booted      = false;
var devMode     = false;            // true when running outside Outlook (local preview)
var savedRange  = null;             // last caret/selection inside the editor
var noteShown   = false;            // one-time deliverability note flag
var _resetPending = false, _resetTimer = null, _previewTimer = null, _statusTimer = null;

// ── Startup ──────────────────────────────────────────────────────────────────

if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady(function (info) {
    boot(!(info && info.host === Office.HostType.Outlook));
  });
}
// Fallback: if Office never initializes (plain browser / blocked CDN), still
// show the editor so it can be previewed and tested locally.
setTimeout(function () { if (!booted) boot(true); }, 1500);

function boot(isDev) {
  if (booted) return;
  booted  = true;
  devMode = isDev;
  wireEditor();
  loadSaved();
  updatePreview();
  document.getElementById('loading').style.display = 'none';
  document.getElementById('app').style.display     = 'flex';
  if (devMode) document.getElementById('devNotice').style.display = 'block';
}

function hasRoaming() {
  try { return !devMode && !!(Office.context && Office.context.roamingSettings); }
  catch (e) { return false; }
}

// ── Editor wiring ────────────────────────────────────────────────────────────

function wireEditor() {
  editor = document.getElementById('editor');

  editor.addEventListener('keyup',  function () { saveRange(); schedulePreview(); });
  editor.addEventListener('mouseup', saveRange);
  editor.addEventListener('input',  schedulePreview);
  editor.addEventListener('paste',  onPaste);

  editor.addEventListener('dragover', function (e) { e.preventDefault(); editor.classList.add('drag-over'); });
  editor.addEventListener('dragleave', function () { editor.classList.remove('drag-over'); });
  editor.addEventListener('drop',     onDrop);

  try { document.execCommand('styleWithCSS', false, true); } catch (e) {}
}

// ── Selection tracking (so toolbar dropdowns can restore the caret) ──────────

function saveRange() {
  var sel = window.getSelection();
  if (sel && sel.rangeCount && editor.contains(sel.anchorNode)) {
    savedRange = sel.getRangeAt(0).cloneRange();
  }
}

function restoreRange() {
  if (!savedRange) return false;
  try {
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(savedRange);
    return true;
  } catch (e) { return false; }
}

function caretToEnd() {
  try {
    var r = document.createRange();
    r.selectNodeContents(editor);
    r.collapse(false);
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
  } catch (e) {}
}

function focusEditor() {
  editor.focus();
  if (!restoreRange()) caretToEnd();
}

// ── Toolbar: simple commands (B / I / U / align / list / clear) ──────────────
// Buttons use onmousedown + preventDefault, so the editor keeps focus and the
// current selection stays live when the command runs.

function cmd(e, command) {
  if (e) e.preventDefault();
  editor.focus();
  try {
    document.execCommand('styleWithCSS', false, true);
    document.execCommand(command, false, null);
  } catch (err) {}
  saveRange();
  schedulePreview();
}

// ── Toolbar: font family ─────────────────────────────────────────────────────

function applyFont(family) {
  if (!family) return;
  focusEditor();
  try {
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('fontName', false, family);
  } catch (e) {}
  document.getElementById('fontSelect').value = '';
  saveRange();
  schedulePreview();
}

// ── Toolbar: font size (px) ──────────────────────────────────────────────────
// execCommand('fontSize') only speaks 1–7. We tag the selection as size 7,
// then rewrite those <font> nodes to inline-styled spans at the chosen px —
// keeping the output email-clean (no <font> tags) and at an exact size.

function applySize(px) {
  if (!px) return;
  focusEditor();
  var sel = window.getSelection();
  if (!sel || sel.isCollapsed) {
    showStatus('Select some text first, then choose a size.', 'info');
    document.getElementById('sizeSelect').value = '';
    return;
  }
  try {
    document.execCommand('styleWithCSS', false, false);   // force <font size=7> markers
    document.execCommand('fontSize', false, '7');
    document.execCommand('styleWithCSS', false, true);
  } catch (e) {}

  toArray(editor.querySelectorAll('font[size="7"]')).forEach(function (f) {
    var span = document.createElement('span');
    span.style.fontSize = px + 'px';
    while (f.firstChild) span.appendChild(f.firstChild);
    f.parentNode.replaceChild(span, f);
  });
  // Defensive: some engines emit a span with xxx-large under styleWithCSS.
  toArray(editor.querySelectorAll('span[style*="xxx-large"]')).forEach(function (s) {
    s.style.fontSize = px + 'px';
  });

  document.getElementById('sizeSelect').value = '';
  saveRange();
  schedulePreview();
}

// ── Toolbar: text color ──────────────────────────────────────────────────────

function applyColor(color) {
  focusEditor();
  try {
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('foreColor', false, color);
  } catch (e) {}
  saveRange();
  schedulePreview();
}

// ── Toolbar: link ────────────────────────────────────────────────────────────

function insertLink(e) {
  if (e) e.preventDefault();
  var url = window.prompt('Link web address:', 'https://');
  if (!url) return;
  url = url.trim();
  if (!/^https?:\/\//i.test(url) && !/^mailto:/i.test(url)) url = 'https://' + url;
  focusEditor();
  try {
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('createLink', false, url);
  } catch (err) {}
  saveRange();
  schedulePreview();
}

// ── Toolbar: image by URL (recommended path for logos) ───────────────────────

function insertImageUrl(e) {
  if (e) e.preventDefault();
  var url = window.prompt('Image web address (https://…):', 'https://');
  if (!url) return;
  url = url.trim();
  if (!/^https:\/\//i.test(url)) {
    showStatus('Please use a secure https:// image address.', 'error');
    return;
  }
  insertHtmlAtCursor('<img src="' + esc(url) + '" style="max-width:100%;height:auto;" alt=""/>');
}

// ── Image paste / drop (embedded base64, size-guarded) ───────────────────────

function onPaste(e) {
  var dt = e.clipboardData;
  if (!dt) return;

  var file = imageFromItems(dt.items);
  if (file) { e.preventDefault(); handleImageFile(file); return; }

  var html = dt.getData && dt.getData('text/html');
  if (html) { e.preventDefault(); insertHtmlAtCursor(sanitizeHtml(html)); }
  // Otherwise let the browser paste plain text normally.
}

function onDrop(e) {
  editor.classList.remove('drag-over');
  var files = e.dataTransfer && e.dataTransfer.files;
  if (!files || !files.length) return;
  for (var i = 0; i < files.length; i++) {
    if (files[i].type && files[i].type.indexOf('image/') === 0) {
      e.preventDefault();
      handleImageFile(files[i]);
      return;
    }
  }
}

function imageFromItems(items) {
  if (!items) return null;
  for (var i = 0; i < items.length; i++) {
    if (items[i].kind === 'file' && items[i].type.indexOf('image/') === 0) {
      return items[i].getAsFile();
    }
  }
  return null;
}

function handleImageFile(file) {
  var reader = new FileReader();
  reader.onload = function () {
    var dataUrl = String(reader.result || '');
    if (dataUrl.length > MAX_IMG_CHARS) {
      showStatus('That image is too large to save with your signature (about ' +
        Math.round(dataUrl.length / 1024) + ' KB; limit ~20 KB). Host it online and add it with the ' +
        'image button (by web address) instead.', 'error');
      return;
    }
    insertHtmlAtCursor('<img src="' + dataUrl + '" style="max-width:100%;height:auto;" alt=""/>');
    deliverabilityNote();
  };
  reader.readAsDataURL(file);
}

function deliverabilityNote() {
  if (noteShown) return;
  noteShown = true;
  showStatus('Image added. Heads-up: pasted images may not display in Gmail and some clients — ' +
    'for a logo, the image button (by web address) is the most reliable.', 'warn');
}

// ── Insert HTML at the caret ─────────────────────────────────────────────────

function insertHtmlAtCursor(html) {
  focusEditor();
  try { document.execCommand('insertHTML', false, html); }
  catch (e) { editor.innerHTML += html; }
  saveRange();
  schedulePreview();
}

// ── Light paste sanitizer (keeps inline styles; drops scripts/classes/mso) ───

function sanitizeHtml(html) {
  var tmp = document.createElement('div');
  tmp.innerHTML = html;

  toArray(tmp.querySelectorAll('script, style, meta, link, title')).forEach(function (n) {
    n.parentNode && n.parentNode.removeChild(n);
  });

  // Remove Office/Word namespaced tags (e.g. <o:p>, <w:*>) by unwrapping them.
  toArray(tmp.getElementsByTagName('*')).forEach(function (el) {
    if (el.tagName && el.tagName.indexOf(':') !== -1) {
      while (el.firstChild) el.parentNode.insertBefore(el.firstChild, el);
      el.parentNode.removeChild(el);
    }
  });

  toArray(tmp.querySelectorAll('*')).forEach(function (el) {
    toArray(el.attributes).forEach(function (attr) {
      var name = attr.name.toLowerCase();
      if (name === 'class' || name === 'id' || name.indexOf('on') === 0 ||
          name === 'lang' || name === 'align' || name.indexOf('xmlns') === 0 ||
          name.indexOf('mso') !== -1 || name.indexOf('data-') === 0) {
        el.removeAttribute(attr.name);
      }
    });
    var style = el.getAttribute && el.getAttribute('style');
    if (style) {
      var cleaned = style.replace(/mso-[^;]+;?/gi, '').replace(/^\s+|\s+$/g, '');
      if (cleaned) el.setAttribute('style', cleaned); else el.removeAttribute('style');
    }
  });

  return tmp.innerHTML;
}

// ── Starter layout ───────────────────────────────────────────────────────────

var STARTER_HTML =
  '<div style="font-family:Arial,Helvetica,sans-serif;">' +
    '<p style="margin:0;font-size:15px;font-weight:bold;color:#000000;">Your Name</p>' +
    '<p style="margin:2px 0 0 0;font-size:12px;color:#555555;">Job Title &nbsp;&bull;&nbsp; Company Name</p>' +
    '<div style="margin:7px 0;height:1px;background-color:#dddddd;line-height:1px;font-size:1px;">&nbsp;</div>' +
    '<p style="margin:0;font-size:12px;color:#333333;">' +
      'Office: 000-000-0000 &nbsp;&nbsp;|&nbsp;&nbsp; ' +
      '<a href="mailto:you@company.com" style="color:#0078d4;text-decoration:none;">you@company.com</a> &nbsp;&nbsp;|&nbsp;&nbsp; ' +
      '<a href="https://company.com" style="color:#0078d4;text-decoration:none;">company.com</a>' +
    '</p>' +
    '<div style="margin:8px 0 0 0;padding:5px 8px;background-color:#f5f5f5;border-left:2px solid #0078d4;font-size:11px;color:#444444;">' +
      'Licensed Professional #00000 &nbsp;&bull;&nbsp; Regulated by [Agency]. Replace with your required compliance text.' +
    '</div>' +
    '<p style="margin:6px 0 0 0;font-size:9px;color:#aaaaaa;line-height:1.4;">' +
      'This email and any attachments are confidential and intended solely for the addressee. ' +
      'If you received this in error, please delete it and notify the sender.' +
    '</p>' +
  '</div>';

function insertStarter() {
  if (editor.innerHTML.trim() &&
      !window.confirm('Replace the current contents with the starter layout?')) {
    return;
  }
  editor.innerHTML = STARTER_HTML;
  caretToEnd();
  saveRange();
  schedulePreview();
  showTab('edit');
  showStatus('Starter layout added — edit or delete any part of it.', 'info');
}

// ── Load / Save (roamingSettings) ────────────────────────────────────────────

function loadSaved() {
  if (!hasRoaming()) return;
  try {
    var html = Office.context.roamingSettings.get(STORAGE_KEY);
    if (typeof html === 'string' && html.length) { editor.innerHTML = html; return; }

    // One-time migration from the old form-field format.
    var legacy = Office.context.roamingSettings.get(LEGACY_KEY);
    if (legacy) {
      var f = typeof legacy === 'string' ? JSON.parse(legacy) : legacy;
      editor.innerHTML = legacyHtml(f);
    }
  } catch (e) { /* first run or corrupt — start blank */ }
}

function saveSignature() {
  if (!hasRoaming()) { showStatus('Save works inside Outlook. (Preview mode.)', 'warn'); return; }

  var html = editor.innerHTML.trim();
  if (!html) { showStatus('Your signature is empty — add something first.', 'error'); return; }
  if (html.length > MAX_SIG_CHARS) {
    showStatus('This signature is too large to save — an embedded image is pushing it over the limit. ' +
      'Add logos with the image button (by web address) instead.', 'error');
    return;
  }

  try {
    Office.context.roamingSettings.set(STORAGE_KEY, html);
    Office.context.roamingSettings.remove(LEGACY_KEY);   // drop superseded form data
    Office.context.roamingSettings.saveAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        showStatus('Saved.', 'success');
      } else {
        var msg = (result.error && result.error.message) || '';
        if (/9057/.test(msg) || /size/i.test(msg)) {
          showStatus('Too large to save — embedded images push it over the limit. ' +
            'Use the image button to add logos by web address.', 'error');
        } else {
          showStatus('Save failed — check your connection and try again.', 'error');
        }
      }
    });
  } catch (e) {
    showStatus('Save failed: ' + e.message, 'error');
  }
}

// ── Insert into the open email ───────────────────────────────────────────────

function insertSignature() {
  if (devMode) { showStatus('Insert works inside Outlook. (Preview mode.)', 'warn'); return; }

  var inner = editor.innerHTML.trim();
  if (!inner) { showStatus('Your signature is empty — add something first.', 'error'); return; }

  var item = (Office.context.mailbox && Office.context.mailbox.item) || null;
  if (!item) { showStatus('No compose window found. Open a new email first.', 'error'); return; }

  var html = wrapSignature(inner) + MARKER;
  var opts = { coercionType: Office.CoercionType.Html };

  if (typeof item.body.setSignatureAsync === 'function') {
    item.body.setSignatureAsync(html, opts, function (r) { insertDone(r, false); });
  } else {
    item.body.prependAsync(html, opts, function (r) { insertDone(r, true); });
  }
}

function insertDone(r, compat) {
  if (r.status === Office.AsyncResultStatus.Succeeded) {
    showStatus(compat ? 'Signature inserted (compatibility mode).' : 'Signature inserted.', 'success');
  } else {
    showStatus('Insert failed: ' + ((r.error && r.error.message) || 'unknown error'), 'error');
  }
}

// Wrap the editor content in a base-font container so the signature has a
// consistent default font even where the user didn't set one. Inline overrides
// inside still win.
function wrapSignature(inner) {
  return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#333333;line-height:1.4;">' +
         inner + '</div>';
}

// ── Reset (double-tap to confirm) ────────────────────────────────────────────

function resetSignature() {
  var btn = document.getElementById('btnReset');

  if (!_resetPending) {
    _resetPending = true;
    btn.textContent = 'Tap again to confirm';
    btn.classList.add('confirming');
    _resetTimer = setTimeout(function () {
      _resetPending = false;
      btn.textContent = 'Reset';
      btn.classList.remove('confirming');
    }, 3000);
    return;
  }

  clearTimeout(_resetTimer);
  _resetPending = false;
  btn.textContent = 'Reset';
  btn.classList.remove('confirming');

  editor.innerHTML = '';
  savedRange = null;
  schedulePreview();

  if (hasRoaming()) {
    Office.context.roamingSettings.remove(STORAGE_KEY);
    Office.context.roamingSettings.remove(LEGACY_KEY);
    Office.context.roamingSettings.saveAsync(function () { showStatus('Signature cleared.', 'success'); });
  } else {
    showStatus('Signature cleared.', 'success');
  }
}

// ── Tabs + preview ───────────────────────────────────────────────────────────

function showTab(name) {
  var tabs = document.querySelectorAll('.tab-btn');
  for (var i = 0; i < tabs.length; i++) {
    tabs[i].classList.toggle('active', tabs[i].dataset.tab === name);
  }
  document.getElementById('tab-edit').style.display    = name === 'edit'    ? 'flex' : 'none';
  document.getElementById('tab-preview').style.display = name === 'preview' ? 'flex' : 'none';
  if (name === 'preview') updatePreview();
}

function schedulePreview() {
  clearTimeout(_previewTimer);
  _previewTimer = setTimeout(updatePreview, 150);
}

function updatePreview() {
  var el = document.getElementById('previewHtml');
  if (!el) return;
  var inner = editor ? editor.innerHTML.trim() : '';
  el.innerHTML = inner
    ? wrapSignature(inner)
    : '<p style="color:#979593;font-size:12px;margin:0;">Your signature preview will appear here.</p>';
}

// ── Legacy migration renderer (old form fields → simple inline HTML) ─────────

function legacyHtml(f) {
  if (!f) return '';
  var accent = f.accentColor || '#0078D4';
  var p = [];
  if (f.fullName) p.push('<p style="margin:0;font-size:15px;font-weight:bold;color:#000;">' + esc(f.fullName) + '</p>');
  var meta = [f.jobTitle, f.company, f.department].filter(Boolean).map(esc);
  if (meta.length) p.push('<p style="margin:2px 0 0 0;font-size:12px;color:#555;">' + meta.join(' &bull; ') + '</p>');
  var c = [];
  if (f.phone)   c.push('Office: ' + esc(f.phone));
  if (f.mobile)  c.push('Mobile: ' + esc(f.mobile));
  if (f.email)   c.push('<a href="mailto:' + esc(f.email) + '" style="color:' + accent + ';text-decoration:none;">' + esc(f.email) + '</a>');
  if (f.website) c.push('<a href="https://' + esc(f.website) + '" style="color:' + accent + ';text-decoration:none;">' + esc(f.website) + '</a>');
  if (c.length)  p.push('<p style="margin:6px 0 0 0;font-size:12px;color:#333;">' + c.join(' &nbsp;|&nbsp; ') + '</p>');
  if (f.address) p.push('<p style="margin:4px 0 0 0;font-size:11px;color:#777;">' + esc(f.address).replace(/\n/g, '<br>') + '</p>');
  if (f.compliance) p.push('<div style="margin:8px 0 0 0;padding:5px 8px;background:#f5f5f5;border-left:2px solid ' + accent + ';font-size:11px;color:#444;">' + esc(f.compliance).replace(/\n/g, '<br>') + '</div>');
  if (f.disclaimer) p.push('<p style="margin:6px 0 0 0;font-size:9px;color:#aaa;">' + esc(f.disclaimer).replace(/\n/g, '<br>') + '</p>');
  return '<div style="font-family:Arial,Helvetica,sans-serif;">' + p.join('') + '</div>';
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function toArray(nodeList) {
  return Array.prototype.slice.call(nodeList || []);
}

function esc(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function showStatus(msg, type) {
  var el = document.getElementById('statusMsg');
  if (!el) return;
  el.textContent  = msg;
  el.className     = 'status-msg status-' + (type || 'info');
  el.style.display = 'block';
  clearTimeout(_statusTimer);
  var dur = (type === 'warn' || type === 'error') ? 7000 : 4000;
  _statusTimer = setTimeout(function () { el.style.display = 'none'; }, dur);
}
