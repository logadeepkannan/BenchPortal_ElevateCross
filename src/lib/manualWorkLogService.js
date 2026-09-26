// Work Log client for the manual (email) sign-in path. Manual sign-in has no
// delegated Graph token, so unlike sharePointService.js (which calls Graph
// directly from the browser), this goes through the small backend in
// /server, which holds the app-only Graph credential — see manualAuth.js for
// why that token exchange can't happen in browser code.

const API_URL = (import.meta.env.VITE_AUTH_API_URL || "").replace(/\/$/, "");

export function isManualWorkLogAvailable() {
  return Boolean(API_URL);
}

async function request(path, options = {}) {
  const res = await fetch(`${API_URL}${path}`, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  const body = res.status === 204 ? null : await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(body?.error || `Request failed (${res.status}).`);
  }
  return body;
}

export function fetchManualProjects() {
  return request("/api/projects");
}

export function createManualProject(project) {
  return request("/api/projects", { method: "POST", body: JSON.stringify(project) });
}

export function updateManualProject(id, project) {
  return request(`/api/projects/${id}`, { method: "PATCH", body: JSON.stringify(project) });
}

export function removeManualProject(id) {
  return request(`/api/projects/${id}`, { method: "DELETE" });
}

export function fetchManualSiteUsers() {
  return request("/api/site-users");
}

export function fetchManualWorkLogs() {
  return request("/api/worklogs");
}

export function createManualWorkLog(log) {
  return request("/api/worklogs", { method: "POST", body: JSON.stringify(log) });
}

export function removeManualWorkLog(id) {
  return request(`/api/worklogs/${id}`, { method: "DELETE" });
}
