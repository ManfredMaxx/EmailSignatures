'use strict';

// ════════════════════════════════════════════════════════════════════════════
//  Signify MOBILE SPIKE — throwaway test handler.
//  Fires on composing a NEW message in Outlook on iOS (OnNewMessageCompose).
//  It answers the one open question for the mobile build:
//    1) Can the mobile event runtime READ roamingSettings written on desktop?
//    2) Does body.setSignatureAsync work on iOS? (the inserted text confirms it)
//  It reports the result in an on-screen notification + the message body.
//  This is a SEPARATE add-in (own GUID) — it does NOT touch production Signify.
// ════════════════════════════════════════════════════════════════════════════

Office.onReady();

function onNewMessageComposeHandler(event) {
  var item = Office.context.mailbox.item;
  var status, inserted;

  try {
    var rs = Office.context.roamingSettings;
    var val = rs && rs.get('spikeTest');
    if (val) {
      status = 'roamingSettings READ OK on iPhone (value: ' + val + ')';
      inserted = '[Signify spike] roamingSettings worked. Read: ' + val;
    } else {
      status = 'roamingSettings was EMPTY — open the spike on desktop, tap "Save test value", then try again.';
      inserted = '[Signify spike] roamingSettings was empty.';
    }
  } catch (e) {
    status = 'roamingSettings ERROR: ' + (e && e.message);
    inserted = '[Signify spike] error reading roamingSettings.';
  }

  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#333333;">' + inserted + '</div>' +
    '<div id="sig-marker" style="display:none;">&#8203;</div>';

  item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, function () {
    try {
      item.notificationMessages.addAsync('spike', {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: ('Signify spike: ' + status).slice(0, 150),
        icon: 'none',
        persistent: false
      }, function () { event.completed(); });
    } catch (e) { event.completed(); }
  });
}

Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
