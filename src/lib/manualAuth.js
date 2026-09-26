// Thin client for the small backend in /server, which checks a manual
// "Sign In" email against the BenchAdministration SharePoint list. This has
// to be a real backend rather than a direct browser call: Microsoft's
// identity platform refuses to issue an app-only (client credentials) Graph
// token to any request carrying a browser Origin header, specifically to
// keep app secrets out of client-side code — so the token exchange has to
// happen server-side regardless. If VITE_AUTH_API_URL isn't set,
// isManualAuthConfigured() is false and Sign In shows a clear "not
// connected" error instead of silently failing.

const API_URL = (import.meta.env.VITE_AUTH_API_URL || "").replace(/\/$/, "");

export function isManualAuthConfigured() {
  return Boolean(API_URL);
}

// Sign-up is paused for now — the server only exposes a login check. An
// admin adds rows to BenchAdministration by hand, or a Microsoft sign-in
// provisions one automatically; anyone whose email is already in that list
// is let in, regardless of password.
export async function verifyManualLogin({ email }) {
  if (!isManualAuthConfigured()) {
    throw new Error("Manual sign-in isn't connected yet — set VITE_AUTH_API_URL in .env to your /server instance's base URL.");
  }
  const res = await fetch(`${API_URL}/api/manual-login`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ email }),
  });
  const body = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(body.error || `Sign-in service error (${res.status}).`);
  }
  return body; // { name, accessLevel }
}
