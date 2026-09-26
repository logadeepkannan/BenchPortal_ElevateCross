import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";

const CLIENT_ID = import.meta.env.VITE_MSAL_CLIENT_ID || "";
const TENANT_ID = import.meta.env.VITE_MSAL_TENANT_ID || "common";

export const GRAPH_SCOPES = ["Sites.ReadWrite.All", "User.Read"];

export function isMsalConfigured() {
  return Boolean(CLIENT_ID);
}

let msalInstance = null;
let initPromise = null;

async function getMsalInstance() {
  if (!isMsalConfigured()) {
    throw new Error("Microsoft sign-in is not configured. Set VITE_MSAL_CLIENT_ID in .env to enable it.");
  }
  if (!msalInstance) {
    msalInstance = new PublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
        redirectUri: window.location.origin,
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
      },
    });
  }
  if (!initPromise) {
    initPromise = msalInstance.initialize();
  }
  await initPromise;
  return msalInstance;
}

// Popup-based sign-in (loginPopup) depends on the opener window reliably
// detecting and closing the popup after redirect — a mechanism that's known
// to be flaky in some browsers (Firefox in particular). When it fails, the
// popup is left open running the whole app a second time inside itself,
// and clicking "Microsoft" again in there throws block_nested_popups. A
// full-page redirect avoids that entire class of bug: there's no second
// window at all, just a navigate-away-and-back. completeMicrosoftSignIn()
// below is what picks up the result after the app reloads.
export async function signInWithMicrosoft() {
  const instance = await getMsalInstance();
  await instance.loginRedirect({ scopes: GRAPH_SCOPES, prompt: "select_account" });
}

// Call once when the app starts. If this load is the app coming back from
// signInWithMicrosoft's redirect, resolves with the signed-in account;
// otherwise (a normal page load) resolves null.
export async function completeMicrosoftSignIn() {
  const instance = await getMsalInstance();
  const result = await instance.handleRedirectPromise();
  if (result?.account) {
    instance.setActiveAccount(result.account);
    return result.account;
  }
  return null;
}

async function acquireToken(scopes) {
  const instance = await getMsalInstance();
  const account = instance.getActiveAccount() || instance.getAllAccounts()[0];
  if (!account) {
    throw new Error("No signed-in Microsoft account. Sign in again.");
  }
  try {
    const result = await instance.acquireTokenSilent({ scopes, account });
    return result.accessToken;
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      // This runs from a background data load, not a click — a popup opened
      // here has no user gesture behind it and browsers routinely block it
      // silently, turning "needs consent" into a confusing dead end. Fail
      // with a clear, actionable message instead of attempting one.
      throw new Error(
        `Microsoft hasn't approved access to "${scopes.join(", ")}" yet. In your Entra app registration, confirm this permission was added ` +
          `and that "Grant admin consent" was clicked for it (adding a permission after an earlier consent needs its own consent click), then sign out and back in.`
      );
    }
    throw err;
  }
}

export async function getGraphAccessToken() {
  return acquireToken(GRAPH_SCOPES);
}

// Microsoft Graph has no equivalent of SharePoint's site user list, which is
// what a Person/Group column's id actually points at. Resolving who a real
// SharePoint "AssignedToLookupId" refers to means calling the SharePoint REST
// API directly — a different resource/audience from Graph, so it needs its
// own token and its own delegated permission (SharePoint API → AllSites.Read).
export async function getSharePointRestToken(hostname) {
  return acquireToken([`https://${hostname}/AllSites.Read`]);
}

export async function signOutMicrosoft() {
  if (!isMsalConfigured() || !msalInstance) return;
  const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  await msalInstance.logoutRedirect({ account });
}
