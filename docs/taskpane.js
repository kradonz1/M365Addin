"use strict";

// ---------------------------------------------------------------------------
// Nested App Authentication (NAA) configuration
// TODO: Replace YOUR_AAD_CLIENT_ID (both below and in manifest.xml / manifest.json)
//       with your Azure AD app registration client ID before deploying.
// The redirect URI registered in Azure AD must include:
//   api://kradonz1.github.io/M365Addin/YOUR_AAD_CLIENT_ID
// ---------------------------------------------------------------------------
var NAA_CLIENT_ID = "YOUR_AAD_CLIENT_ID";
var NAA_AUTHORITY = "https://login.microsoftonline.com/common";
var NAA_SCOPES = ["openid", "profile", "User.Read"];
var SAFETY_EMAIL = "AresSafetyCoalition-SM@aresmgmt.com";

var msalInstance = null;

/**
 * Initialize MSAL using Nested App Authentication (NAA).
 * NAA lets the Office host broker tokens on behalf of the add-in so that
 * sign-in dialogs are never shown to the user.
 * Falls back to the standard PublicClientApplication when NAA is unavailable.
 */
async function initMsal() {
  if (NAA_CLIENT_ID === "YOUR_AAD_CLIENT_ID") {
    console.error(
      "NAA: NAA_CLIENT_ID is still a placeholder. " +
      "Replace YOUR_AAD_CLIENT_ID in taskpane.js, manifest.xml, and manifest.json " +
      "with your Azure AD app registration client ID before deploying."
    );
  }
  var msalConfig = {
    auth: {
      clientId: NAA_CLIENT_ID,
      authority: NAA_AUTHORITY
    },
    system: {
      allowNativeBroker: false
    }
  };

  if (typeof msal !== "undefined" && typeof msal.createNestablePublicClientApplication === "function") {
    try {
      msalInstance = await msal.createNestablePublicClientApplication(msalConfig);
      console.log("NAA: initialized via createNestablePublicClientApplication");
      return;
    } catch (e) {
      console.warn("NAA: createNestablePublicClientApplication failed, falling back:", e);
    }
  }

  if (typeof msal !== "undefined") {
    msalInstance = new msal.PublicClientApplication(msalConfig);
    await msalInstance.initialize();
    console.log("NAA: initialized via PublicClientApplication (fallback)");
  }
}

/**
 * Acquire an access token.  Attempts a silent request first; falls back to a
 * popup only when the host cannot broker the token silently.
 */
async function acquireToken() {
  if (!msalInstance) return null;

  var tokenRequest = { scopes: NAA_SCOPES };
  var accounts = msalInstance.getAllAccounts();

  if (accounts.length > 0) {
    try {
      var result = await msalInstance.acquireTokenSilent(
        Object.assign({}, tokenRequest, { account: accounts[0] })
      );
      console.log("NAA: token acquired silently");
      return result;
    } catch (silentError) {
      console.warn("NAA: silent acquisition failed:", silentError.message);
    }
  }

  try {
    var popupResult = await msalInstance.acquireTokenPopup(tokenRequest);
    console.log("NAA: token acquired via popup");
    return popupResult;
  } catch (popupError) {
    console.error("NAA: popup acquisition failed:", popupError.message);
    return null;
  }
}

/** Show the signed-in user's display name and UPN in the task pane. */
function displayUserInfo(tokenResponse) {
  var el = document.getElementById("user-info");
  if (!el || !tokenResponse || !tokenResponse.account) return;
  var account = tokenResponse.account;
  var displayName = account.name || account.username || "Unknown user";
  var upn = account.username || "";

  // Build DOM nodes to avoid XSS from token response values
  while (el.firstChild) el.removeChild(el.firstChild);
  el.appendChild(document.createTextNode("\u2713 Signed in as: "));
  var strong = document.createElement("strong");
  strong.textContent = displayName;
  el.appendChild(strong);
  if (upn) {
    el.appendChild(document.createTextNode(" (" + upn + ")"));
  }
  el.style.display = "block";
}

/** Resolve the Outlook Web App (OWA) origin for deep-link compose URLs. */
function getOwaOrigin() {
  var restUrl = Office.context && Office.context.mailbox ? Office.context.mailbox.restUrl : null;
  if (restUrl && /^https?:\/\//i.test(restUrl)) {
    return new URL(restUrl).origin;
  }
  try {
    var topHref = window.top && window.top.location ? window.top.location.href : null;
    if (topHref && /^https?:\/\//i.test(topHref)) {
      return new URL(topHref).origin;
    }
  } catch (e) {}
  return "https://outlook.office.com";
}

/** Open a pre-filled compose window addressed to the safety coalition. */
function openPrefilledCompose() {
  var hostName = Office.context &&
    Office.context.mailbox &&
    Office.context.mailbox.diagnostics
      ? Office.context.mailbox.diagnostics.hostName
      : null;

  if (hostName !== "OutlookWebApp") {
    window.location.href = "mailto:" + encodeURIComponent(SAFETY_EMAIL);
  } else {
    var origin = getOwaOrigin();
    var url = origin + "/mail/deeplink/compose" +
      "?to=" + encodeURIComponent(SAFETY_EMAIL) +
      "&subject=" + encodeURIComponent("") +
      "&body=" + encodeURIComponent("");
    window.open(url, "_blank", "noopener,noreferrer");
  }
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------
Office.onReady(async function (info) {
  if (info.host === Office.HostType.Outlook) {
    var diagnostics = Office.context.mailbox.diagnostics;
    console.log("UPN:", diagnostics);
    console.log("HostName:", diagnostics.hostName);
    console.log("OWA version:", diagnostics.hostVersion);

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Authenticate with NAA
    try {
      await initMsal();
      var tokenResponse = await acquireToken();
      if (tokenResponse) {
        displayUserInfo(tokenResponse);
      }
    } catch (authError) {
      console.error("NAA: authentication failed:", authError);
    }

    document.getElementById("composeEmergencyNews").onclick = openPrefilledCompose;
  }
});

