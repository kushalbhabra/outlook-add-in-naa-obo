# OneDrive / SharePoint File Picker v8 — Integration Guide

Research notes for embedding the **Microsoft-hosted File Picker v8** inside an **Outlook add-in** (and Teams apps). The picker lets users open, save, and share files/folders stored in OneDrive and SharePoint using the exact same UI as the M365 service — no custom file UI needed.

> Source of truth: [OneDrive/samples](https://github.com/OneDrive/samples) → `samples/file-picking/` and the [official docs](https://aka.ms/OneDrive/file-picker).

---

## 1. What it is

The File Picker v8 is **not an SDK you import**. It is a **page hosted inside the Microsoft service** (`/_layouts/15/FilePicker.aspx`) that you load in an **iframe or popup** and drive over `postMessage` + `MessageChannel`. Because it's service-hosted, every improvement Microsoft ships to the real picker automatically appears in your app.

Key facts:

- Works with **OneDrive for Business**, **SharePoint sites**, and **OneDrive Consumer (personal)**.
- Always uses **delegated permissions** — it can only access files/folders the signed-in user can already see.
- Supports **files**, **folders**, and **save-as** flows (see §4).
- Requires **SharePoint tokens** (not Graph tokens) for the `authenticate` command.

---

## 2. How it works (protocol)

```
1. POST a form  →  {baseUrl}/_layouts/15/FilePicker.aspx?filePicker=<JSON config>
                   (hidden field: access_token = <SharePoint token>)
2. Picker sends  →  { type: "initialize", channelId }  to the host window
3. Host grabs    →  event.ports[0]  (a MessagePort), port.start(), send { type: "activate" }
4. Over the port →  handle commands: "authenticate" | "pick" | "close"
```

`{baseUrl}` is either the user's OneDrive (`https://{tenant}-my.sharepoint.com`) or a SharePoint web URL (`https://{tenant}.sharepoint.com/sites/dev`). For OneDrive Consumer use `https://onedrive.live.com/picker`.

### Minimal configuration

```ts
const options = {
  sdk: "8.0",
  entry: { oneDrive: { files: {} } },          // or sharePoint: { files: {} }
  authentication: {},                          // REQUIRED for iframe embedding / full item data
  messaging: {
    origin: "https://addin-domain.com",        // host app origin
    channelId: crypto.randomUUID(),            // always a fresh GUID per instance
  },
  typesAndSources: {
    mode: "files",                             // "files" | "folders" | "all"
    pivots: { oneDrive: true, recent: true },
  },
};
```

### Message listener (canonical)

```ts
let port: MessagePort;

window.addEventListener("message", (event) => {
  if (event.source !== pickerWindow) return;              // pickerWindow = iframe.contentWindow or popup
  const message = event.data;
  if (message.type === "initialize" && message.channelId === options.messaging.channelId) {
    port = event.ports[0];
    port.addEventListener("message", channelMessageListener);
    port.start();
    port.postMessage({ type: "activate" });
  }
});

async function channelMessageListener(message: MessageEvent) {
  const payload = message.data;
  if (payload.type === "notification") {
    if (payload.data?.notification === "page-loaded") { /* picker ready */ }
    return;
  }
  if (payload.type !== "command") return;

  port.postMessage({ type: "acknowledge", id: payload.id });   // ALL commands must be acked
  const command = payload.data;

  switch (command.command) {
    case "authenticate": {
      // command.resource = the baseUrl the token must be for
      const token = await getSharePointToken(command.resource);   // see §5
      port.postMessage({ type: "result", id: payload.id,
        data: { result: "token", token } });
      break;
    }
    case "pick": {
      const items = command.items;                               // picked files/folders
      handlePicked(items);                                       // e.g. messageParent() back to task pane
      port.postMessage({ type: "result", id: payload.id, data: { result: "success" } });
      break;
    }
    case "close":
      closePicker();
      break;
    default:
      port.postMessage({ type: "result", id: payload.id,
        data: { result: "error", error: { code: "unsupportedCommand", message: command.command } } });
  }
}
```

### Picked item shape (always guaranteed)

```ts
{
  id: string,                       // driveItem id
  parentReference: { driveId: string },
  "@sharePoint.endpoint": string,   // e.g. https://{tenant}-my.sharepoint.com/_api/v2.1
}
```

Build a Graph URL: `@sharePoint.endpoint + /drives/ + parentReference.driveId + /items/ + id`. Folders are the same shape with a `folder: { childCount }` facet.

---

## 3. Embedding in an Outlook add-in

Outlook add-ins run inside an **iframe** (web) or **webview** (desktop), so `window.open()` is unreliable. **Always use the Office Dialog API** (`Office.context.ui.displayDialogAsync`) or an in-task-pane iframe.

### Option A — Picker in an Office dialog (recommended, more space)

The picker's recommended size is **1080×680**, too big for a task pane. Open a dialog that hosts the picker iframe. The dialog page **must be on the same domain** as the add-in.

**Task pane (`taskpane.html`):**

```ts
Office.onReady(() => {
  document.getElementById("pick").onclick = () => {
    Office.context.ui.displayDialogAsync(
      "https://addin-domain.com/picker.html",          // same domain as add-in
      { height: 70, width: 60, displayInIframe: true },
      (asyncResult) => {
        const dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          const picked = JSON.parse(arg.message);      // items from the picker
          console.log(picked);
          dialog.close();
        });
      }
    );
  };
});
```

**Dialog page (`picker.html`)** — hosts the picker iframe and runs the §2 protocol:

```ts
const url = `${baseUrl}/_layouts/15/FilePicker.aspx?filePicker=${encodeURIComponent(JSON.stringify(options))}`;

const frame = document.getElementById("frame");        // <iframe sandbox="allow-same-origin allow-forms allow-scripts">
const form = frame.contentWindow.document.createElement("form");
form.setAttribute("action", url);
form.setAttribute("method", "POST");
const input = document.createElement("input");         // hidden, name="access_token", value=<SharePoint token>
form.appendChild(input);
frame.contentWindow.document.body.append(form);
form.submit();

// ... then the §2 message listener, with:
//   "pick"   → Office.context.ui.messageParent(JSON.stringify(command.items))
//   "close"  → Office.context.ui.messageParent("close")
```

### Option B — Picker embedded directly in the task pane

Same messaging code, but the iframe lives in the task pane itself. Simpler, but cramped — only for small pickers.

---

## 4. Folder support

Folder picking is **purely a config change** — the integration code is identical.

| Scenario | Config |
|---|---|
| **Folder-only picker** | `typesAndSources: { mode: "folders", filters: ["folder"] }` |
| **Start inside a folder** (OneDrive) | `entry: { oneDrive: { files: { folder: "Pictures" } } }` |
| **Start inside a folder** (SharePoint) | `entry: { sharePoint: { byPath: { web, list, folder } } }` |
| **Save-As dialog** | `tray: { prompt: "save-as", saveAs: { fileName } }` + `mode: "folders"` |
| **Restrict source location** | `typesAndSources.locations.sharePoint.byPath.folder` |
| **Allow create/upload** | `commands.createFolder.enabled`, `commands.upload.enabled` |

Multi-select: `selection: { mode: "multiple" }`. Pre-select items: `selection.sourceItems`.

---

## 5. Authentication

The picker needs **SharePoint tokens** for `command.resource` (the baseUrl). Request scopes `[`${resource}/.default`]`. Three patterns:

### A. MSAL.js in the add-in (simplest)
`acquireTokenSilent` / `loginPopup` with scopes `[`${resource}/.default`]`. Works everywhere but may prompt.

### B. Office SSO + backend OBO (best UX — matches this repo)
1. Add-in: `Office.auth.getAccessToken()` → send to your backend.
2. Backend: exchange via **OBO** for a SharePoint token (same pattern as `app-b/` in this repo, but target resource = the SharePoint baseUrl instead of Graph).
3. Return the token to the picker's `authenticate` command.

### C. NAA (Nested App Authentication)
`acquireTokenSilent(["api://<app-id>/access_as_user"])` then OBO in the backend — the exact chain this repo demonstrates. Requires the `brk-multihub://` redirect URI and self-pre-authorization.

> The picker relies on **SharePoint tokens, not Graph** — you cannot reuse a Graph token for the picker.

### Required AAD permissions (delegated)

| Operation | OneDrive | SharePoint Sites |
|---|---|---|
| Read | `SharePoint.MyFiles.Read` or `Graph.Files.Read` | `SharePoint.MyFiles.Read` / `Graph.Files.Read` / `SharePoint.AllSites.Read` |
| Write | `SharePoint.MyFiles.Write` or `Graph.Files.ReadWrite` | `SharePoint.MyFiles.Write` / `Graph.Files.ReadWrite` / `SharePoint.AllSites.Write` |

---

## 6. Manifest changes (Outlook add-in)

```json
{
  "webApplicationInfo": {
    "id": "<AAD app client id>",
    "resource": "https://addin-domain.com/<AAD app client id>"
  },
  "AppDomains": [
    { "Domain": "*.sharepoint.com" },
    { "Domain": "*.my.sharepoint.com" },
    { "Domain": "onedrive.live.com" }
  ]
}
```

`AppDomains` is what lets the iframe/dialog load the picker page. The dialog page must be same-domain as the add-in.

---

## 7. Teams apps (from `OneDrive/samples` → `teams-picker`)

The repo's `teams-picker` is a **Teams Toolkit personal tab** that runs in Teams, Outlook, and the M365 app:

1. **Tab** opens the picker with the Teams dialog API:
   ```ts
   dialog.url.open({
     url: "https://localhost:53000/#/picker-dialog",
     size: { height: 800, width: 1000 },
     title: "File Picker",
   }, (result) => console.log(result));
   ```
2. **`PickerDialog.tsx`** hosts the picker iframe + §2 protocol, but auth is TeamsFx:
   ```ts
   const apiClient = createApiClient(`${functionEndpoint}/api/`,
     new BearerTokenAuthProvider(async () => (await teamsUserCredential.getToken(""))!.token));
   ```
3. **`api/getOBOToken/index.ts`** — Azure Function that exchanges the Teams user token for a SharePoint token via `OnBehalfOfUserCredential` (OBO) — same backend pattern as §5B.
4. On `pick`, returns items via `dialog.url.submit(JSON.stringify(command))`.

---

## 8. Reference samples (in `OneDrive/samples`)

| Sample | Use when |
|---|---|
| `samples/file-picking/javascript-basic/` | **Canonical** — no SDK, plain JS + MSAL. Start here. |
| `samples/file-picking/typescript-react/` | React + `@pnp/picker-api` wrapper (`Picker(...).using(Popup(), LamdaAuthenticate(getToken))`). |
| `samples/file-picking/teams-picker/` | Teams Toolkit personal tab (Teams/Outlook/M365). |
| `samples/file-picking/example-picker-configs/` | Ready-made configs: save-as, folder-path, search, OneDrive consumer, SharePoint pages. |
| `samples/file-browsing/` | The non-picking **File Browser** control (browse-only). |

---

## 9. Gotchas

- **`window.open` is unreliable in Office add-ins** — use `displayDialogAsync` (or `openBrowserWindow` for external URLs).
- **Dialog page must be same full domain** as the add-in host page.
- **`access_token` hidden field is required** when embedding in an iframe (optional for popups).
- **Always ack every command** before responding, and always reply (even `unsupportedCommand`).
- **Fresh `channelId` GUID** per picker instance.
- **Large `sourceItems`** can exceed URL/POST size limits — keep pre-selections small.
- **`authentication: {}` must be present** to get full item data / embed in an iframe.
- Picker works with **either OneDrive OR SharePoint** per instance — include only one `entry` section.