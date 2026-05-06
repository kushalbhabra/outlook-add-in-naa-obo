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
