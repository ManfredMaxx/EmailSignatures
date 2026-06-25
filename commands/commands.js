'use strict';

// ════════════════════════════════════════════════════════════════════════════
//  Signify — send guard (OnMessageSend)
//
//  R0 / PRIME DIRECTIVE: fail CLOSED. If we cannot positively confirm a
//  signature is present, we BLOCK or PROMPT — we never silently allow an
//  unsigned message to send. A conscious "Send Anyway" (express permission)
//  is offered only when a signature is genuinely missing.
//
//  Platform note: the manifest SendMode is the ultimate backstop for when this
//  handler can't run at all (error loading / offline). "Block" is the only mode
//  that refuses to send in that case — see REQUIREMENTS.md (R0).
// ════════════════════════════════════════════════════════════════════════════

// Initialize the event runtime. Missing this is a common cause of the handler
// never completing in new Outlook on Windows (the "processing… / taking longer"
// dialog that never resolves).
Office.onReady();

function onMessageSend(event) {
  var settled = false;

  // Complete exactly once.
  function settle(options) {
    if (settled) return;
    settled = true;
    clearTimeout(guard);
    event.completed(options);
  }

  // Safety net: never hang. If we don't have an answer quickly, FAIL CLOSED
  // (block, don't allow). Kept under Outlook's own ~5s long-running threshold
  // so the user sees our clear message instead of the generic stuck dialog.
  var guard = setTimeout(function () {
    settle({
      allowEvent: false,
      errorMessage:
        'Signify couldn’t verify your signature in time. Open Signify, ' +
        'click “Insert into this email,” then send again.'
    });
  }, 4000);

  try {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          // Couldn't read the body — FAIL CLOSED (do NOT allow).
          settle({
            allowEvent: false,
            errorMessage:
              'Signify couldn’t check your message for a signature. Open ' +
              'Signify, click “Insert into this email,” then send again.'
          });
          return;
        }

        var body = result.value || '';
        var hasSignature = body.indexOf('sig-marker') !== -1;

        if (hasSignature) {
          settle({ allowEvent: true });
          return;
        }

        // Signature genuinely missing. Block, but offer a conscious override
        // ("Send Anyway" = express permission). The runtime send-mode override
        // (requirement set 1.14) makes this work even under a "Block" manifest;
        // if it isn't supported, we fall back to the manifest send mode.
        var options = {
          allowEvent: false,
          errorMessage:
            'No signature found. Open Signify and click “Insert into this ' +
            'email” before sending. To send without one, choose “Send Anyway.”'
        };
        try {
          var SMO = Office.MailboxEnums && Office.MailboxEnums.SendModeOverride;
          if (SMO && SMO.PromptUser) options.sendModeOverride = SMO.PromptUser;
        } catch (e) { /* override unsupported on this client — use manifest mode */ }

        settle(options);
      }
    );
  } catch (err) {
    // Unexpected failure — FAIL CLOSED.
    settle({
      allowEvent: false,
      errorMessage:
        'Signify hit an error checking your message. Open Signify, click ' +
        '“Insert into this email,” then send again.'
    });
  }
}

// Register the handler (must match manifest <LaunchEvent FunctionName="...">).
Office.actions.associate('onMessageSend', onMessageSend);
