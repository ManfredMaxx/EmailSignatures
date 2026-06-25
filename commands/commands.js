'use strict';

// ════════════════════════════════════════════════════════════════════════════
//  Signify — send guard (OnMessageSend)
//
//  R0 / PRIME DIRECTIVE: fail CLOSED. If we cannot positively confirm a
//  signature is present, we BLOCK or PROMPT — we never silently allow an
//  unsigned message to send. A conscious "Send Anyway" (express permission)
//  is offered only when a signature is genuinely missing.
//
//  The one exception is the user's own standing opt-out: a prominent,
//  off-by-default, warning-gated toggle in the task pane sets
//  roamingSettings['sendBlockDisabled'] = true. Honouring that is itself
//  express permission (the user consciously turned the check off).
//
//  Platform note: the manifest SendMode is the ultimate backstop for when this
//  handler can't run at all (error loading / offline). "Block" is the only mode
//  that refuses to send in that case — see REQUIREMENTS.md (R0).
// ════════════════════════════════════════════════════════════════════════════

Office.onReady();

function onMessageSend(event) {
  var settled = false;

  var guard = setTimeout(function () {
    settle({
      allowEvent: false,
      errorMessage:
        'Signify couldn’t verify your signature in time. Open Signify, ' +
        'click “Insert into this email,” then send again.'
    });
  }, 4000);

  function settle(options) {
    if (settled) return;
    settled = true;
    clearTimeout(guard);
    event.completed(options);
  }

  // Standing user opt-out (conscious, warned, off by default).
  try {
    var rs = Office.context.mailbox && Office.context.roamingSettings;
    if (rs && rs.get('sendBlockDisabled') === true) {
      settle({ allowEvent: true });
      return;
    }
  } catch (e) { /* can't read setting — fall through to the fail-closed check */ }

  try {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          settle({
            allowEvent: false,
            errorMessage:
              'Signify couldn’t check your message for a signature. Open ' +
              'Signify, click “Insert into this email,” then send again.'
          });
          return;
        }

        var body = result.value || '';
        if (body.indexOf('sig-marker') !== -1) {
          settle({ allowEvent: true });
          return;
        }

        var options = {
          allowEvent: false,
          errorMessage:
            'No signature found. Open Signify and click “Insert into this ' +
            'email” before sending. To send without one, choose “Send Anyway.”'
        };
        try {
          var SMO = Office.MailboxEnums && Office.MailboxEnums.SendModeOverride;
          if (SMO && SMO.PromptUser) options.sendModeOverride = SMO.PromptUser;
        } catch (e) { /* override unsupported — use manifest send mode */ }

        settle(options);
      }
    );
  } catch (err) {
    settle({
      allowEvent: false,
      errorMessage:
        'Signify hit an error checking your message. Open Signify, click ' +
        '“Insert into this email,” then send again.'
    });
  }
}

Office.actions.associate('onMessageSend', onMessageSend);
