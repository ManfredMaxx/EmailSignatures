'use strict';

// Fires when the user clicks Send.
// Reads the email body and checks for the hidden sig-marker element.
// If missing, blocks send and shows a message. If present, allows send.
function onMessageSend(event) {
  Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Html,
    function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        // Can't read the body — fail open so a network hiccup doesn't permanently block send
        event.completed({ allowEvent: true });
        return;
      }

      var body = result.value || '';
      var hasSignature = body.indexOf('sig-marker') !== -1;

      if (hasSignature) {
        event.completed({ allowEvent: true });
      } else {
        event.completed({
          allowEvent: false,
          errorMessage:
            'No signature found. Click "Email Signature" in the ribbon, ' +
            'then click "Insert into this email" before sending.'
        });
      }
    }
  );
}

// Binds the function name declared in manifest.xml <LaunchEvent FunctionName="...">
Office.actions.associate('onMessageSend', onMessageSend);
