/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, console */

import { AccountManager } from "./authConfig";
import { makeGraphRequest } from "./msgraph-helper";

const accountManager = new AccountManager();
const sideloadMsg = document.getElementById("sideload-msg");
const appBody = document.getElementById("app-body");
const getUserDataButton = document.getElementById("getUserData");
const getUserFilesButton = document.getElementById("getUserFiles");
const getDebugTokenButton = document.getElementById("getDebugToken");
const getRegATokenButton = document.getElementById("getRegAToken");
const getAppBTokenButton = document.getElementById("getAppBToken");
const readMailAppBButton = document.getElementById("readMailAppB");
const userName = document.getElementById("userName");
const userEmail = document.getElementById("userEmail");

// Initialize when Office is ready.
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";
    if (getUserDataButton) {
      getUserDataButton.addEventListener("click", getUserData);
    }
    if (getUserFilesButton) {
      getUserFilesButton.addEventListener("click", getUserFiles);
    }
    if (getDebugTokenButton) {
      getDebugTokenButton.addEventListener("click", showDebugToken);
    }
    if (getRegATokenButton) {
      getRegATokenButton.addEventListener("click", showRegAToken);
    }
    if (getAppBTokenButton) {
      getAppBTokenButton.addEventListener("click", showAppBToken);
    }
    if (readMailAppBButton) {
      readMailAppBButton.addEventListener("click", readMailViaAppB);
    }
    // Initialize MSAL.
    accountManager.initialize();
  }
});

/**
 * Writes a list of filenames into the email body.
 * @param fileNameList The list of filenames.
 */
async function writeFileNames(fileNameList: string[]) {
  const item = Office.context.mailbox.item;
  let fileNameBody: string = "";
  fileNameList.map((fileName) => (fileNameBody += "<br/>" + fileName));

  if (item) {
    item.body.setAsync(fileNameBody, {
      coercionType: "html",
    });
  }
}

/**
 * Gets the user data such as name and email and displays it
 * in the task pane.
 */
async function getUserData() {
  const userDataElement = document.getElementById("userData");
  // Specify minimum scopes for the token needed.
  const accessToken = await accountManager.ssoGetAccessToken(["user.read"]);
  console.log("[DEBUG] Access token (getUserData):", accessToken);

  const response: { displayName: string; mail: string } = await makeGraphRequest(accessToken, "/me", "");

  if (userDataElement) {
    userDataElement.style.visibility = "visible";
  }
  if (userName) {
    userName.innerText = response.displayName ?? "";
  }
  if (userEmail) {
    userEmail.innerText = response.mail ?? "";
  }
}

/**
 * Gets the first 10 item names (files or folders) from the user's OneDrive.
 * Inserts the item names into the document.
 */
async function getUserFiles() {
  try {
    const names = await getFileNames(10);

    writeFileNames(names);
  } catch (error) {
    console.error(error);
  }
}

/**
 * Gets item names (files or folders) from the user's OneDrive.
 */
async function getFileNames(count = 10) {
  // Specify minimum scopes for the token needed.
  const accessToken = await accountManager.ssoGetAccessToken(["Files.Read"]);
  console.log("[DEBUG] Access token (getFileNames):", accessToken);
  const response: { value: { name: string }[] } = await makeGraphRequest(
    accessToken,
    "/me/drive/root/children",
    `?$select=name&$top=${count}`
  );

  const names = response.value.map((item: { name: string }) => item.name);
  return names;
}

/**
 * Reads the user's recent emails by calling App B's /api/mail endpoint.
 * App B validates the token and does OBO → Graph on the server side.
 */
async function readMailViaAppB() {
  const mailList = document.getElementById("mailList");
  const mailItems = document.getElementById("mailItems");

  if (mailItems) mailItems.innerHTML = "<em>Loading…</em>";
  if (mailList) mailList.style.display = "block";

  try {
    const token = await accountManager.ssoGetAccessToken([
      "api://outlook-target-api/access_as_user",
    ]);

    const response = await fetch("http://localhost:4000/api/mail", {
      headers: { Authorization: `Bearer ${token}` },
    });

    const data = await response.json() as {
      error?: string;
      messages?: { subject: string; from: { emailAddress: { name: string } }; receivedDateTime: string; isRead: boolean }[];
    };

    if (!response.ok || data.error) {
      if (mailItems) mailItems.innerHTML = `<span style="color:red">Error: ${data.error}</span>`;
      return;
    }

    if (mailItems) {
      mailItems.innerHTML = (data.messages ?? []).map((m) =>
        `<div style="padding:5px 0; border-bottom:1px solid #eee">
          <div style="font-weight:${m.isRead ? "400" : "700"}">${m.subject || "(no subject)"}</div>
          <div style="color:#666; font-size:10px">${m.from?.emailAddress?.name} &middot; ${new Date(m.receivedDateTime).toLocaleDateString()}</div>
        </div>`
      ).join("");
    }
  } catch (error) {
    if (mailItems) mailItems.innerHTML = `<span style="color:red">${JSON.stringify(error, Object.getOwnPropertyNames(error))}</span>`;
  }
}

/**
 * DEBUG: Acquires a token scoped directly to App B via NAA (aud = App B client ID).
 * Proves the taskpane can call App B without any middle-tier server.
 * Requires Reg A's client ID pre-authorized in App B's Expose an API settings.
 */
async function showAppBToken() {
  const debugTokenArea = document.getElementById("debugTokenArea");
  const debugTokenEl = document.getElementById("debugToken") as HTMLTextAreaElement;
  try {
    const token = await accountManager.ssoGetAccessToken(["api://outlook-target-api/access_as_user"]);
    console.log("[DEBUG] App B token (direct NAA):", token);
    if (debugTokenArea) debugTokenArea.style.display = "block";
    if (debugTokenEl) debugTokenEl.value = token;
  } catch (error) {
    if (debugTokenArea) debugTokenArea.style.display = "block";
    if (debugTokenEl) debugTokenEl.value = JSON.stringify(error, Object.getOwnPropertyNames(error), 2);
  }
}

/**
 * DEBUG: Gets a token for Reg A scope via NAA (aud = Reg A client ID).
 * Use this token as the OBO assertion. Requires "api://<client-id>/access_as_user"
 * to be exposed on Reg A and admin consent granted.
 */
async function showRegAToken() {
  const debugTokenArea = document.getElementById("debugTokenArea");
  const debugTokenEl = document.getElementById("debugToken") as HTMLTextAreaElement;
  try {
    const token = await accountManager.ssoGetAccessToken(["api://00e7ef2d-f553-4ffa-9076-616b18eb04c5/access_as_user"]);
    console.log("[DEBUG] Reg A token (OBO assertion):", token);
    if (debugTokenArea) debugTokenArea.style.display = "block";
    if (debugTokenEl) debugTokenEl.value = token;
  } catch (error) {
    if (debugTokenArea) debugTokenArea.style.display = "block";
    if (debugTokenEl) debugTokenEl.value = JSON.stringify(error, Object.getOwnPropertyNames(error), 2);
  }
}

/**
 * DEBUG: Gets the raw Office SSO token and displays it in the task pane.
 * This uses Office.auth.getAccessToken() which returns a token with aud = Reg A client ID,
 * suitable as the OBO assertion. NAA cannot be used here since the Office broker only
 * issues tokens for Microsoft-owned resources (Graph, Exchange).
 */
async function showDebugToken() {
  const debugTokenArea = document.getElementById("debugTokenArea");
  const debugTokenEl = document.getElementById("debugToken") as HTMLTextAreaElement;
  try {
    const token = await Office.auth.getAccessToken({ allowSignInPrompt: true });
    if (debugTokenArea) debugTokenArea.style.display = "block";
    if (debugTokenEl) debugTokenEl.value = token;
  } catch (error) {
    if (debugTokenArea) debugTokenArea.style.display = "block";
    if (debugTokenEl) debugTokenEl.value = JSON.stringify(error, Object.getOwnPropertyNames(error), 2);
  }
}
