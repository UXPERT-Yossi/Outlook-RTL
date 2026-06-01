# RTL Default — Outlook patch

Makes **New Outlook for Mac** compose every message right-to-left: new emails, replies,
reply-all, and forwards. Implemented as a Microsoft-supported **event-based Outlook
add-in** (no SIP hacks, no app patching — survives Outlook updates).

> **Why an add-in and not a "real" patch?** New Outlook for Mac is app-sandboxed and
> code-signed by Microsoft, and macOS SIP blocks injecting code into it. The only
> supported, update-proof way to change compose behaviour is the add-in extensibility
> model, which writes to the message body via Office.js on the `OnMessageCompose` event.

---

## What's in here

| File | Purpose |
|------|---------|
| `manifest.xml` | The add-in definition. Declares the `OnMessageCompose` launch event. |
| `commands.html` | Headless runtime page that loads Office.js + the handler. |
| `src/launchevent.js` | The handler. **You implement `buildRtlBody()`** — see step 1. |
| `assets/icon-*.png` | Icons referenced by the manifest. |
| `configure.sh` | Bakes your GitHub Pages URL into the manifest. |
| `.nojekyll` | Tells GitHub Pages to serve `src/` and dotfiles as-is. |

---

## Step 1 — Finish the handler (required)

Open `src/launchevent.js` and implement `buildRtlBody(currentHtml)`. It currently
returns the body unchanged (no-op). Pick a strategy from the comments in that file:

- **Strategy A (full wrap)** — most reliable RTL for your typing; flips quoted text too.
- **Strategy B (prepend RTL block)** — gentlest on replies; quoted text stays as-is.

Keep the `RTL_MARKER` string on the wrapper element you add, or the idempotency guard
won't recognise it and the body will be re-wrapped on every compose event.

## Step 2 — Host the files on GitHub Pages

1. Create a new **public** GitHub repo, e.g. `outlook-rtl-patch`.
2. Push these files to it:
   ```bash
   git init
   git add -A
   git commit -m "RTL Default Outlook add-in"
   git branch -M main
   git remote add origin https://github.com/<you>/outlook-rtl-patch.git
   git push -u origin main
   ```
3. In the repo: **Settings → Pages → Build and deployment → Source: Deploy from a
   branch**, branch `main`, folder `/ (root)`. Save.
4. Wait ~1 minute, then confirm `https://<you>.github.io/outlook-rtl-patch/commands.html`
   loads (blank page = correct).

## Step 3 — Point the manifest at your URL

```bash
./configure.sh <your-github-username> outlook-rtl-patch
```
This replaces every `__BASEURL__` in `manifest.xml`. Commit and push the updated manifest.

## Step 4 — Deploy the add-in (no admin required)

Event-based add-ins auto-launch when **admin-deployed** OR when **sideloaded**. You are
not a tenant admin, so use the self-service sideload path below. (Sideloading is
officially a testing mechanism — it works per-user and auto-launches the
`OnMessageCompose` event, but the mailbox may clear it periodically, so you might need
to re-add it occasionally.)

### Sideload to your own mailbox
1. Download your configured manifest: open
   `https://<you>.github.io/outlook-rtl-patch/manifest.xml` in a browser and save the
   file locally. (The old "Add from URL" option has been removed, so you must add from
   a local file.)
2. Go to <https://aka.ms/olksideload>. Outlook on the web opens and, after a few
   seconds, the **Add-Ins for Outlook** dialog appears.
3. Select **My add-ins** → under **Custom Addins**, choose **Add a custom add-in →
   Add from File**.
4. Pick the `manifest.xml` you downloaded and accept all prompts.
5. The add-in syncs to New Outlook for Mac for the same mailbox. Restart the app.

### Optional: ask IT to deploy it for everyone
If you'd rather have a persistent, no-maintenance install, send your admin the manifest
URL and ask them to deploy it via **admin.microsoft.com → Settings → Integrated apps →
Upload custom apps → Provide link to manifest file**. Then you can skip sideloading.

## Step 5 — Test

Open New Outlook for Mac → **New message**. The body should start RTL (right-aligned,
cursor on the right). Try a **reply** and a **forward** too. If nothing happens, see
Troubleshooting.

---

## Path B — Make the *reading pane* RTL too

The add-in only controls **compose**. The read surface is locked down for security, so
incoming mail direction is a mailbox setting. Hebrew/Arabic content already auto-renders
RTL, but to flip the reading pane and UI fully:

1. Open **outlook.office.com** → **Settings (gear) → General → Language and time**.
2. Set **Language** to a right-to-left language (e.g. **עברית / Hebrew**), **or** leave
   the language and tick **"Display in right-to-left orientation"** if shown.
3. Save. This setting lives on the mailbox server and **syncs down to New Outlook for
   Mac** automatically (restart the app to pick it up).

> Trade-off: switching the display language to Hebrew also flips the whole Outlook UI
> (menus, panels) to RTL and translates labels. If you want English menus but RTL
> reading, the RTL-orientation checkbox (when available for your account) is the lever;
> otherwise the language switch is the only built-in option.

---

## Troubleshooting

- **Nothing happens on compose.** Confirm the manifest URL loads, that `configure.sh`
  ran (no `__BASEURL__` left: `grep __BASEURL__ manifest.xml` should be empty), and that
  the add-in still appears under **My add-ins → Custom Addins** (sideloaded add-ins can
  be cleared by the mailbox — re-add it if it's gone). Give New Outlook for Mac a restart.
- **Body unchanged.** You haven't implemented `buildRtlBody()` yet (Step 1).
- **Patch runs once but not after that.** Expected if the `RTL_MARKER` is present — the
  guard skips already-RTL bodies. Remove the marker to force re-apply while testing.
- **Add-in won't load at all.** Your client must support Mailbox requirement set **1.12**.
  Update New Outlook for Mac.
- **Debugging:** errors are logged via `console.error`. See Microsoft's
  "Troubleshoot event-based and spam-reporting add-ins" guide for attaching a debugger.

## Updating later

Edit files → push to GitHub. The web/Mac runtime loads JS fresh from Pages, so handler
changes apply on next compose (admin-deployed manifest *structure* changes need
re-consent in the admin center; pure `launchevent.js` logic changes do not).
