---
page_type: sample
urlFragment: outlook-add-in-sso-naa
products:
  - m365
  - office
  - office-outlook
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: "03/19/2024 10:00:00 AM"
description: "Outlook add-in SSO with NAA, plus an App B backend that reads email via Microsoft Graph on behalf of the user."
---

# Outlook add-in with SSO using nested app authentication

## Summary

This sample demonstrates **NAA (Nested App Authentication)** in an Outlook add-in using MSAL.js. It extends the standard SSO sample with a second API backend (**App B**) that receives a delegated token from the taskpane and reads the signed-in user's mail via Microsoft Graph — with no middle-tier server on the add-in side.

### What's included

- **Outlook add-in** (`src/`) — NAA-based taskpane with a "Read Mail" button that calls App B directly.
- **App B server** (`app-b/`) — Hono backend that validates the incoming token, performs an OBO exchange for a Graph token, and returns the user's 10 most recent emails.

### Auth chain

```
[Outlook taskpane]
    │  NAA: acquireTokenSilent(["api://<app-b-id>/access_as_user"])
    ▼
[App B — http://localhost:4000]
    │  1. Validate Bearer token (aud = App B, scp = access_as_user)
    │  2. OBO → Graph token (acquireTokenOnBehalfOf)
    ▼
[Microsoft Graph — /me/messages]
```

## Features

- MSAL.js NAA for silent SSO — no sign-in prompt in the normal case.
- Direct browser → App B token acquisition (no App A server relay).
- App B validates the token with AAD JWKS and exchanges it for a Graph token via OBO.
- Fallback to Office dialog API when NAA is unavailable.

## Applies to

- Outlook on Windows (new and classic), Mac, mobile, and on the web.

## Prerequisites

- Office connected to a Microsoft 365 subscription.
- [Node.js](https://nodejs.org/) v18 or later.
- [npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm) v8 or greater.
- Access to the [Azure portal](https://portal.azure.com) to create app registrations.

## Setup

### Step 1 — Register App A (the Outlook add-in)

1. Go to [Azure portal → App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) → **New registration**.
2. Set **Name** to `Outlook-Add-in-SSO-NAA`, **Supported account types** to single-tenant.
3. Add these **SPA redirect URIs** (Authentication blade):
   - `brk-multihub://localhost:3000` ← required for NAA broker
   - `https://localhost:3000/auth.html`
   - `https://localhost:3000/dialog.html`
4. Under **Expose an API**:
   - Set Application ID URI to `api://<app-a-client-id>`.
   - Add scope `access_as_user` (Admins and users, Enabled).
   - Under **Authorized client applications**, add App A's **own client ID** with `access_as_user` selected.  
     _(Self-pre-authorization is required for NAA silent token acquisition to work.)_
5. In the **Manifest** editor, inside the `api` object set `"requestedAccessTokenVersion": 2`.
6. Copy the **Application (client) ID**.

### Step 2 — Register App B (the backend)

1. Create a new registration named `Outlook-Target-API`, **single-tenant**.
2. Under **Expose an API**:
   - Set Application ID URI to `api://<app-b-client-id>`.
   - Add scope `access_as_user` (Admins and users, Enabled).
   - Under **Authorized client applications**, add **App A's client ID** with `access_as_user`.  
     This lets the Outlook taskpane acquire an App B token via NAA without a consent prompt.
3. In the **Manifest** editor set `"accessTokenAcceptedVersion": 2`.
4. Under **API permissions** → add delegated Microsoft Graph permissions:
   - `Mail.Read`
   - `User.Read`
   - `Calendars.Read` (optional)
   - Click **Grant admin consent**.
5. Under **Certificates & secrets** → **New client secret** → save the value.

### Step 3 — Configure the add-in

1. Open `src/taskpane/msalconfig.ts` and replace `"Enter_the_Application_Id_Here"` with App A's client ID.
2. Open `src/taskpane/taskpane.ts` and replace the App B scope placeholder with `api://<app-b-client-id>/access_as_user`.

### Step 4 — Configure App B

Create `app-b/.env` (already git-ignored):

```
APP_B_CLIENT_ID=<app-b-client-id>
APP_B_CLIENT_SECRET=<app-b-client-secret>
TENANT_ID=<your-tenant-id>
```

## Run the sample

Open **two terminals**.

**Terminal 1 — Outlook add-in:**

```bash
npm install
npm run start
```

This starts the webpack dev server at `https://localhost:3000` and sideloads the add-in into Outlook.

**Terminal 2 — App B server:**

```bash
cd app-b
npm install
npm run start
```

This starts the Hono server at `http://localhost:4000`.

**In Outlook:**

1. Open or compose an email and click **Show task pane** on the ribbon.
2. Click **Read Mail (via App B)** — the taskpane acquires an App B token via NAA, sends it to `http://localhost:4000/api/mail`, and displays your 10 most recent emails.

## Manifest types

By default the sample uses the add-in only manifest. To switch to the unified manifest, copy all files from `manifest-configurations/unified/` to the root (and delete `manifest.xml`). To switch back, copy from `manifest-configurations/add-in-only/`.

## How it works

### Add-in (`src/`)

- `src/taskpane/authConfig.ts` — `AccountManager` wraps MSAL NAA. `ssoGetAccessToken(scopes)` acquires tokens silently via the Office broker; falls back to `acquireTokenPopup` then the Office dialog API.
- `src/taskpane/taskpane.ts` — UI logic. The **Read Mail (via App B)** button:
  1. Calls `ssoGetAccessToken(["api://<app-b-id>/access_as_user"])` to get an App B token via NAA.
  2. POSTs to `http://localhost:4000/api/mail` with the token as a Bearer header.
  3. Renders the returned email list in the taskpane.
- `src/taskpane/msgraph-helper.ts` — helper for direct Graph REST calls (used by other buttons).

### App B (`app-b/`)

| File | Purpose |
|---|---|
| `app-b/src/index.ts` | Hono server. CORS for `https://localhost:3000`. Routes `GET /api/mail` and `GET /health`. |
| `app-b/src/validateToken.ts` | Verifies Bearer token with AAD JWKS (`jose`). Checks `aud` (bare GUID or `api://` URI), `iss`, `scp = access_as_user`. |
| `app-b/src/oboHelper.ts` | Singleton `ConfidentialClientApplication`. OBO exchange: App B token → Graph token with `Mail.Read` + `User.Read`. |

### Fallback authentication

If NAA is unavailable (older Outlook clients), `acquireTokenPopup` is tried. If the popup fails with `popup_window_error`, the Office dialog API opens a full auth dialog instead. See `src/taskpane/fallback/fallbackauthdialog.ts`.

## Debugging

Open the project in VS Code, select **Run and Debug** (Ctrl+Shift+D), choose **Outlook Desktop (Edge Chromium)** from the dropdown, and press F5.

See [Debug Office Add-ins on Windows using VS Code and Edge WebView2](https://learn.microsoft.com/office/dev/add-ins/testing/debug-desktop-using-edge-chromium) for more detail.

## Troubleshooting

### NAA token acquisition errors (taskpane → App B)

**Error code `3399614475` — "Access denied for the resource"**

The scope URI in the taskpane doesn't match App B's `identifierUris`. The NAA broker compares them exactly.

- Check App B's manifest: `"identifierUris": ["api://..."]`
- The scope passed to `ssoGetAccessToken` must be `<identifierUri>/access_as_user`
- If App B uses a custom name (e.g. `api://outlook-target-api`), use that — not the GUID form `api://<client-id>`

**Error code `3399614466` (`ff000000`) — NAA silent acquisition failed**

App A's own client ID is not pre-authorized on itself.

- Azure portal → App A → Expose an API → Authorized client applications
- Add App A's **own** client ID with the `access_as_user` scope ticked
- Without this, the NAA broker cannot issue tokens silently

### OBO exchange errors (App B → Graph)

**`AADSTS70003` — invalid_grant / unsupported grant type**

The OBO grant type URN is wrong. This is the correct value:
```
urn:ietf:params:oauth:grant-type:jwt-bearer
```
There is **no `2`** after `oauth`. Using `PowerShell`'s `Invoke-RestMethod` with a hashtable body silently corrupts the colon-delimited value — use `curl.exe --data-urlencode` instead.

**`AADSTS65001` — consent required**

Graph delegated permissions on App B haven't been admin-consented.

- Azure portal → App B → API permissions → Grant admin consent for \<tenant\>
- Required permissions: `Mail.Read`, `User.Read` (and any others your app uses)

### App B token validation errors

**`unexpected 'aud' claim value`**

`jose`'s `jwtVerify` rejected the token because the `aud` claim is a bare GUID (`40cd1669-...`) rather than the `api://` URI form.

Accept both in `validateToken.ts`:
```typescript
audience: [APP_B_CLIENT_ID, `api://${APP_B_CLIENT_ID}`, "api://outlook-target-api"]
```

**`AADSTS70011` — invalid scope**

The `access_as_user` scope doesn't exist or is disabled on App B.

- Azure portal → App B → Expose an API → check the scope is listed and **Enabled**

### Testing the OBO exchange manually

Use `curl.exe` (not `Invoke-RestMethod`) to avoid PowerShell encoding issues:

```powershell
$assertion = "<app-b-token-from-taskpane>"
curl.exe -s -X POST `
  "https://login.microsoftonline.com/<tenant-id>/oauth2/v2.0/token" `
  --data-urlencode "grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer" `
  --data-urlencode "client_id=<app-b-client-id>" `
  --data-urlencode "client_secret=<app-b-client-secret>" `
  --data-urlencode "assertion=$assertion" `
  --data-urlencode "scope=https://graph.microsoft.com/.default" `
  --data-urlencode "requested_token_use=on_behalf_of"
```

Decode the resulting token at [jwt.ms](https://jwt.ms) and confirm `aud = https://graph.microsoft.com`.

## Security reporting

If you find a security issue with our libraries or services, report the issue to [secure@microsoft.com](mailto:secure@microsoft.com). For more information, see [Microsoft Security Response Center](https://technet.microsoft.com/security/dd252948).

## More resources

- NAA docs: https://aka.ms/NAAdocs
- NAA FAQ: https://aka.ms/NAAFAQ
- NAA Word/Excel/PowerPoint sample: https://aka.ms/NAAsampleOffice
- [Register an application with Microsoft Identity Platform](https://learn.microsoft.com/graph/auth-register-app-v2)

## Questions and feedback

- [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) for problems with the sample.
- [Office samples survey](https://aka.ms/OfficeSamplesSurvey) for feedback and suggestions.
- [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) — tag: `office-js-dev`.

## Copyright

Copyright (c) 2024 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/outlook-add-in-sso-naa" />
