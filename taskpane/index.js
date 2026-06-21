'use strict';

// ── Constants ────────────────────────────────────────────────────────────────

var STORAGE_KEY   = 'sigFields';
var FIELD_IDS     = ['fullName','jobTitle','company','department','phone',
                     'mobile','email','website','linkedin','address',
                     'compliance','disclaimer'];
var DEFAULT_COLOR = '#0078D4';

// ── Startup ──────────────────────────────────────────────────────────────────

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    loadSavedFields();
    updatePreview();
    document.getElementById('loading').style.display = 'none';
    document.getElementById('app').style.display     = 'flex';
  }
});

// ── Tab navigation ───────────────────────────────────────────────────────────

function showTab(name) {
  document.querySelectorAll('.tab-btn').forEach(function (btn) {
    btn.classList.toggle('active', btn.dataset.tab === name);
  });
  document.getElementById('tab-edit').style.display    = name === 'edit'    ? '' : 'none';
  document.getElementById('tab-preview').style.display = name === 'preview' ? '' : 'none';
  if (name === 'preview') updatePreview();
}

// ── Roaming settings ─────────────────────────────────────────────────────────

function loadSavedFields() {
  try {
    var raw = Office.context.roamingSettings.get(STORAGE_KEY);
    if (raw) fillForm(typeof raw === 'string' ? JSON.parse(raw) : raw);
  } catch (e) { /* first run or corrupt data — start blank */ }
}

function fillForm(fields) {
  FIELD_IDS.forEach(function (id) {
    var el = document.getElementById(id);
    if (el && fields[id] !== undefined) el.value = fields[id];
  });
  var colorEl = document.getElementById('accentColor');
  if (colorEl && fields.accentColor) colorEl.value = fields.accentColor;
}

function getFormFields() {
  var f = { accentColor: getVal('accentColor') || DEFAULT_COLOR };
  FIELD_IDS.forEach(function (id) { f[id] = getVal(id); });
  return f;
}

function getVal(id) {
  var el = document.getElementById(id);
  return el ? el.value.trim() : '';
}

// ── Save ─────────────────────────────────────────────────────────────────────

function saveSignature() {
  var fields = getFormFields();
  if (!fields.fullName) { showStatus('Full name is required.', 'error'); return; }

  Office.context.roamingSettings.set(STORAGE_KEY, fields);
  Office.context.roamingSettings.saveAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      showStatus('Saved.', 'success');
    } else {
      showStatus('Save failed — check your connection and try again.', 'error');
    }
  });
}

// ── Insert ───────────────────────────────────────────────────────────────────

function insertSignature() {
  var fields = getFormFields();
  if (!fields.fullName) { showStatus('Please fill in at least your full name.', 'error'); return; }

  var item = Office.context.mailbox && Office.context.mailbox.item;
  if (!item) { showStatus('No compose window found. Open a new email first.', 'error'); return; }

  var html = generateHtml(fields);

  if (typeof item.body.setSignatureAsync === 'function') {
    item.body.setSignatureAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          showStatus('Signature inserted.', 'success');
        } else {
          showStatus('Insert failed: ' + result.error.message, 'error');
        }
      }
    );
  } else {
    // Fallback for clients that don't yet support setSignatureAsync
    item.body.prependAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          showStatus('Signature inserted (compatibility mode).', 'success');
        } else {
          showStatus('Insert failed: ' + result.error.message, 'error');
        }
      }
    );
  }
}

// ── Reset ────────────────────────────────────────────────────────────────────

var _resetPending = false;
var _resetTimer   = null;

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

  Office.context.roamingSettings.remove(STORAGE_KEY);
  Office.context.roamingSettings.saveAsync(function () {
    FIELD_IDS.forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.value = '';
    });
    document.getElementById('accentColor').value = DEFAULT_COLOR;
    updatePreview();
    showStatus('Signature cleared.', 'success');
  });
}

// ── Preview ───────────────────────────────────────────────────────────────────

var _previewTimer = null;

function schedulePreview() {
  clearTimeout(_previewTimer);
  _previewTimer = setTimeout(updatePreview, 150);
}

function updatePreview() {
  var fields = getFormFields();
  var el = document.getElementById('previewHtml');
  if (!el) return;
  el.innerHTML = fields.fullName
    ? generateHtml(fields)
    : '<p style="color:#979593;font-size:12px;margin:0;">Fill in your name to see a preview.</p>';
}

// ── HTML generator ────────────────────────────────────────────────────────────

function esc(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escBr(str) {
  return esc(str).replace(/\n/g, '<br>');
}

function toUrl(val, fallbackPrefix) {
  if (!val) return '';
  val = val.trim();
  if (/^https?:\/\//i.test(val)) return val;
  if (/^\/\//i.test(val)) return 'https:' + val;
  return fallbackPrefix + val;
}

function generateHtml(f) {
  var accent = f.accentColor || DEFAULT_COLOR;
  var rows = [];

  // Name
  rows.push(
    '<p style="margin:0;padding:0;font-size:15px;font-weight:bold;color:#000000;line-height:1.3;">' +
    esc(f.fullName) + '</p>'
  );

  // Title • Company • Department
  var meta = [f.jobTitle, f.company, f.department].filter(Boolean).map(esc);
  if (meta.length) {
    rows.push(
      '<p style="margin:2px 0 0 0;padding:0;font-size:12px;color:#555555;line-height:1.4;">' +
      meta.join(' &nbsp;&bull;&nbsp; ') + '</p>'
    );
  }

  // Horizontal rule
  rows.push('<div style="margin:7px 0 5px 0;height:1px;background-color:#DDDDDD;"></div>');

  // Contact details
  var contacts = [];
  if (f.phone) {
    contacts.push(
      '<span style="color:' + accent + ';font-weight:bold;">P</span>&nbsp;' +
      '<a href="tel:' + esc(f.phone) + '" style="color:#333333;text-decoration:none;">' + esc(f.phone) + '</a>'
    );
  }
  if (f.mobile) {
    contacts.push(
      '<span style="color:' + accent + ';font-weight:bold;">M</span>&nbsp;' +
      '<a href="tel:' + esc(f.mobile) + '" style="color:#333333;text-decoration:none;">' + esc(f.mobile) + '</a>'
    );
  }
  if (f.email) {
    contacts.push(
      '<a href="mailto:' + esc(f.email) + '" style="color:' + accent + ';text-decoration:none;">' + esc(f.email) + '</a>'
    );
  }
  if (f.website) {
    var webHref = toUrl(f.website, 'https://');
    contacts.push(
      '<a href="' + esc(webHref) + '" style="color:' + accent + ';text-decoration:none;">' + esc(f.website) + '</a>'
    );
  }
  if (contacts.length) {
    rows.push(
      '<p style="margin:0;padding:0;font-size:12px;line-height:1.8;">' +
      contacts.join('&nbsp;&nbsp;|&nbsp;&nbsp;') + '</p>'
    );
  }

  // LinkedIn
  if (f.linkedin) {
    var liHref = toUrl(f.linkedin, 'https://linkedin.com/in/');
    if (f.linkedin.indexOf('linkedin.com') !== -1 && !/^https?:\/\//i.test(f.linkedin)) {
      liHref = 'https://' + f.linkedin;
    }
    rows.push(
      '<p style="margin:1px 0 0 0;padding:0;font-size:11px;">' +
      '<a href="' + esc(liHref) + '" style="color:' + accent + ';text-decoration:none;">LinkedIn</a>' +
      '</p>'
    );
  }

  // Address
  if (f.address) {
    rows.push(
      '<p style="margin:2px 0 0 0;padding:0;font-size:11px;color:#777777;line-height:1.5;">' +
      escBr(f.address) + '</p>'
    );
  }

  // Compliance block
  if (f.compliance) {
    rows.push(
      '<div style="margin:8px 0 0 0;padding:5px 8px;background-color:#F5F5F5;' +
      'border-left:2px solid ' + accent + ';font-size:11px;color:#444444;line-height:1.5;">' +
      escBr(f.compliance) + '</div>'
    );
  }

  // Legal disclaimer
  if (f.disclaimer) {
    rows.push(
      '<p style="margin:6px 0 0 0;padding:0;font-size:9px;color:#AAAAAA;line-height:1.4;max-width:480px;">' +
      escBr(f.disclaimer) + '</p>'
    );
  }

  return [
    '<table cellpadding="0" cellspacing="0" border="0"',
    ' style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#333333;border-collapse:collapse;">',
    '<tr><td style="border-left:3px solid ' + accent + ';padding:2px 0 2px 12px;vertical-align:top;">',
    rows.join('\n'),
    '</td></tr></table>',
    '<div id="sig-marker" style="display:none;height:0;overflow:hidden;"></div>'
  ].join('');
}

// ── Status messages ───────────────────────────────────────────────────────────

var _statusTimer = null;

function showStatus(msg, type) {
  var el = document.getElementById('statusMsg');
  el.textContent = msg;
  el.className = 'status-msg status-' + (type || 'info');
  el.style.display = 'block';
  clearTimeout(_statusTimer);
  _statusTimer = setTimeout(function () { el.style.display = 'none'; }, 4000);
}
