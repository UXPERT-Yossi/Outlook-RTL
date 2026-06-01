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

function onMessageComposeHandler(event) {
  var item = Office.context.mailbox.item;

  item.body.getAsync(Office.CoercionType.Html, function (getResult) {
    if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("RTL Default: getAsync failed: " + JSON.stringify(getResult.error));
      event.completed();
      return;
    }

    var currentHtml = getResult.value || "";

    // Idempotency guard: if we've already applied RTL, do nothing.
    if (currentHtml.indexOf(RTL_MARKER) !== -1) {
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
  // Strategy B: prepend an empty RTL paragraph where the cursor lands, and leave
  // any quoted reply/forward content below in its original direction.
  var rtlBlock = '<div dir="rtl" style="text-align:right;" ' + RTL_MARKER + '><br></div>';
  return rtlBlock + currentHtml;
}

// REQUIRED: map the manifest's FunctionName ("onMessageComposeHandler") to the JS function.
// Without this association the handler never runs.
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
