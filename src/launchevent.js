/*
 * RTL Default — Outlook event-based add-in
 *
 * Fires on OnMessageCompose (new message, reply, reply all, forward, draft edit)
 * and forces the compose editor to render right-to-left.
 *
 * Runtime notes (these surprised me, so worth writing down):
 *  - Office.onReady / Office.initialize do NOT run in the event-based runtime.
 *    Put all startup logic inside the handler itself.
 *  - You MUST call event.completed() or the runtime hangs for ~300s then is killed.
 *  - UI APIs (dialogs, displayNewMessageForm, item.close) are blocked here.
 *    Reading/writing item.body IS allowed — that's all we need.
 */

// A marker so re-entrant compose events (e.g. reopening a draft) don't double-wrap.
var RTL_MARKER = "data-rtl-default=\"1\"";

// TEMPORARY diagnostic notification so we can see whether the handler runs.
// Remove once RTL is confirmed working.
function notify(item, text) {
  try {
    item.notificationMessages.addAsync("rtlDefaultDiag", {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: ("RTL Default: " + text).substring(0, 150)
    });
  } catch (e) {
    console.error("RTL Default: notify threw: " + e);
  }
}

function onMessageComposeHandler(event) {
  var item = Office.context.mailbox.item;
  console.log("RTL Default: handler fired");
  notify(item, "handler fired");

  item.body.getAsync(Office.CoercionType.Html, function (getResult) {
    if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("RTL Default: getAsync failed: " + JSON.stringify(getResult.error));
      notify(item, "getAsync FAILED");
      event.completed();
      return;
    }

    var currentHtml = getResult.value || "";

    // Idempotency guard: if we've already applied RTL, do nothing.
    if (currentHtml.indexOf(RTL_MARKER) !== -1) {
      console.log("RTL Default: already applied, skipping");
      notify(item, "already applied, skipped");
      event.completed();
      return;
    }

    var newHtml = buildRtlBody(currentHtml);

    item.body.setAsync(
      newHtml,
      { coercionType: Office.CoercionType.Html },
      function (setResult) {
        if (setResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("RTL Default: setAsync failed: " + JSON.stringify(setResult.error));
          notify(item, "setAsync FAILED: " + JSON.stringify(setResult.error));
        } else {
          console.log("RTL Default: body set RTL");
          notify(item, "applied RTL OK");
        }
        // Always signal completion, success or not.
        event.completed();
      }
    );
  });
}

/**
 * buildRtlBody — decide HOW to make the compose body render right-to-left.
 *
 * @param {string} currentHtml  The existing body HTML. For a brand-new message this
 *                              is usually empty or just your signature. For a reply or
 *                              forward it contains your signature PLUS the quoted
 *                              original message ("On <date>, X wrote: ...").
 * @returns {string}            The HTML to write back. MUST include the string in
 *                              RTL_MARKER somewhere on the wrapper element, otherwise
 *                              the idempotency guard above won't recognise our work
 *                              and the body will be re-wrapped every time the event fires.
 *
 * ── The decision (pick one; each is ~3–8 lines) ──────────────────────────────
 *
 *  Strategy A — Full wrap (simplest, most aggressive):
 *      Wrap everything in one RTL div. Quoted reply text flips RTL too.
 *      return '<div dir="rtl" style="text-align:right;" ' + RTL_MARKER + '>' + currentHtml + '</div>';
 *      Tradeoff: an English original you're replying to gets force-aligned right,
 *      which can look odd. But your *new* typing is reliably RTL.
 *
 *  Strategy B — Prepend an empty RTL block (gentlest on replies):
 *      Put a fresh RTL paragraph at the top where your cursor lands, leave the
 *      quoted text below untouched in its original direction.
 *      return '<div dir="rtl" style="text-align:right;" ' + RTL_MARKER + '><br></div>' + currentHtml;
 *      Tradeoff: cleaner for mixed-language threads, but the quoted block stays LTR.
 *
 *  Strategy C — Your own hybrid:
 *      e.g. wrap only when currentHtml is "small" (new message), else prepend.
 *      You have the raw HTML — you decide.
 *
 *  Whichever you choose, keep RTL_MARKER on the outer element you add.
 * ─────────────────────────────────────────────────────────────────────────────
 */
function buildRtlBody(currentHtml) {
  // Wrap the ENTIRE body in one RTL container (so the caret is RTL wherever Outlook
  // places it) and include a leading empty line, also inside the container, so a
  // brand-new message starts with the cursor on the right.
  //
  // Note: this also flips quoted reply/forward text to RTL. We can refine to preserve
  // the original direction once we've confirmed the cursor lands correctly.
  return '<div dir="rtl" style="text-align:right;" ' + RTL_MARKER + '>' +
         '<div><br></div>' + currentHtml +
         '</div>';
}

// REQUIRED: map the manifest's FunctionName ("onMessageComposeHandler") to the JS function.
// Without this association the handler never runs.
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
