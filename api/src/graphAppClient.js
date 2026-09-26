// Server-side only: acquires an app-only (client credentials) Microsoft Graph
// token using a confidential client secret. This must never run in browser
// code — GRAPH_CLIENT_SECRET grants write access to the whole SharePoint site
// it's scoped to, tenant-wide if the app registration was granted
// Sites.ReadWrite.All rather than Sites.Selected.
const { ConfidentialClientApplication } = require("@azure/msal-node");

const TENANT_ID = process.env.GRAPH_TENANT_ID;
const CLIENT_ID = process.env.GRAPH_CLIENT_ID;
const CLIENT_SECRET = process.env.GRAPH_CLIENT_SECRET;
const SP_HOSTNAME = process.env.SP_HOSTNAME;
const SP_SITE_PATH = process.env.SP_SITE_PATH;

function isGraphAppConfigured() {
  return Boolean(TENANT_ID && CLIENT_ID && CLIENT_SECRET && SP_HOSTNAME && SP_SITE_PATH);
}

let msalApp = null;
function getMsalApp() {
  if (!msalApp) {
    msalApp = new ConfidentialClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
        clientSecret: CLIENT_SECRET,
      },
    });
  }
  return msalApp;
}

async function getAppToken() {
  const result = await getMsalApp().acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  return result.accessToken;
}

async function graphFetch(token, path, options = {}) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Graph API ${res.status} ${res.statusText}: ${body.slice(0, 300)}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

async function resolveSiteId(token) {
  const path = SP_SITE_PATH.startsWith("/") ? SP_SITE_PATH : `/${SP_SITE_PATH}`;
  const site = await graphFetch(token, `/sites/${SP_HOSTNAME}:${path}`);
  return site.id;
}

async function resolveListId(token, siteId, displayName) {
  const result = await graphFetch(token, `/sites/${siteId}/lists?$filter=displayName eq '${displayName}'`);
  const list = result.value && result.value[0];
  if (!list) throw new Error(`SharePoint list "${displayName}" was not found on the configured site.`);
  return list.id;
}

// Mirrors the client-side resolution in src/lib/sharePointService.js — lists
// created through the modern Microsoft Lists app often store columns under
// generic internal names that don't match what you typed as the display name.
async function getColumnNameMap(token, siteId, listId) {
  const result = await graphFetch(token, `/sites/${siteId}/lists/${listId}/columns`);
  const map = new Map();
  (result.value || []).forEach((col) => {
    if (col.displayName) map.set(col.displayName, col.name);
  });
  return map;
}

module.exports = {
  isGraphAppConfigured,
  getAppToken,
  graphFetch,
  resolveSiteId,
  resolveListId,
  getColumnNameMap,
};
