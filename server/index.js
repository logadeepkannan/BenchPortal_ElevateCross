require("dotenv").config();
const express = require("express");
const cors = require("cors");
const {
  isGraphAppConfigured,
  getAppToken,
  graphFetch,
  resolveSiteId,
  resolveListId,
  getColumnNameMap,
} = require("./graphAppClient");
const {
  fetchProjects,
  createProject,
  updateProject,
  removeProject,
  fetchSiteUsers,
  fetchWorkLogs,
  createWorkLog,
  removeWorkLog,
} = require("./workLogsService");

const LIST_NAME = "BenchAdministration";
const ACCESS_LEVELS = ["Admin", "Manager", "User"];
const DEFAULT_ACCESS_LEVEL = "User";
const PORT = process.env.PORT || 8787;
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || "*";

const app = express();
app.use(cors({ origin: ALLOWED_ORIGIN }));
app.use(express.json());

// Manual (email-only, for now) Sign In: allows anyone whose email already
// has a row in BenchAdministration. Sign-up isn't wired up yet — an admin
// adds rows by hand, or a Microsoft sign-in provisions one automatically.
app.post("/api/manual-login", async (req, res) => {
  if (!isGraphAppConfigured()) {
    return res.status(501).json({
      error: "This server isn't configured yet — set GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, SP_HOSTNAME, SP_SITE_PATH.",
    });
  }

  const email = String(req.body?.email || "").trim();
  if (!email) {
    return res.status(400).json({ error: "email is required." });
  }

  try {
    const token = await getAppToken();
    const siteId = await resolveSiteId(token);
    const listId = await resolveListId(token, siteId, LIST_NAME);
    const columnMap = await getColumnNameMap(token, siteId, listId);
    const nameField = columnMap.get("UserName") || "UserName";
    const emailField = columnMap.get("UserEmail") || "UserEmail";
    const levelField = columnMap.get("Access_Level") || "Access_Level";

    const items = await graphFetch(token, `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500`);
    const normalized = email.toLowerCase();
    const match = (items.value || []).find(
      (item) => String(item.fields?.[emailField] || "").trim().toLowerCase() === normalized
    );
    if (!match) {
      return res.status(404).json({ error: "This email isn't in BenchAdministration yet. Ask an admin to add you." });
    }

    const level = match.fields?.[levelField];
    return res.json({
      name: match.fields?.[nameField] || email,
      accessLevel: ACCESS_LEVELS.includes(level) ? level : DEFAULT_ACCESS_LEVEL,
    });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || "Failed to check BenchAdministration." });
  }
});

const NOT_CONFIGURED_ERROR = {
  error: "This server isn't configured yet — set GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, SP_HOSTNAME, SP_SITE_PATH.",
};

// Work Log endpoints for the manual (email) sign-in path — mirrors what
// src/lib/sharePointService.js does with a delegated MSAL token, but through
// this server's app-only credential, since manual sign-in has no Graph token
// of its own to call SharePoint directly from the browser with.
app.get("/api/projects", async (req, res) => {
  if (!isGraphAppConfigured()) return res.status(501).json(NOT_CONFIGURED_ERROR);
  try {
    return res.json(await fetchProjects());
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || "Failed to load projects." });
  }
});

app.post("/api/projects", async (req, res) => {
  if (!isGraphAppConfigured()) return res.status(501).json(NOT_CONFIGURED_ERROR);
  const title = String(req.body?.title || "").trim();
  if (!title) return res.status(400).json({ error: "title is required." });
  try {
    return res.status(201).json(await createProject({ ...req.body, title }));
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || "Failed to create the project." });
  }
});

app.patch("/api/projects/:id", async (req, res) => {
  if (!isGraphAppConfigured()) return res.status(501).json(NOT_CONFIGURED_ERROR);
  const title = String(req.body?.title || "").trim();
  if (!title) return res.status(400).json({ error: "title is required." });
  try {
    return res.json(await updateProject(req.params.id, { ...req.body, title }));
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || "Failed to update the project." });
  }
});

app.delete("/api/projects/:id", async (req, res) => {
  if (!isGraphAppConfigured()) return res.status(501).json(NOT_CONFIGURED_ERROR);
  try {
    await removeProject(req.params.id);
    return res.status(204).end();
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || "Failed to delete the project." });
  }
});

app.get("/api/site-users", async (req, res) => {
  if (!isGraphAppConfigured()) return res.status(501).json(NOT_CONFIGURED_ERROR);
  try {
    return res.json(await fetchSiteUsers());
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || "Failed to load site users." });
  }
});

app.get("/api/worklogs", async (req, res) => {
  if (!isGraphAppConfigured()) return res.status(501).json(NOT_CONFIGURED_ERROR);
  try {
    return res.json(await fetchWorkLogs());
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || "Failed to load work logs." });
  }
});

app.post("/api/worklogs", async (req, res) => {
  if (!isGraphAppConfigured()) return res.status(501).json(NOT_CONFIGURED_ERROR);
  const { taskTitle, project, hours, category, status, commit, date, submittedBy } = req.body || {};
  if (!taskTitle || !project) {
    return res.status(400).json({ error: "taskTitle and project are required." });
  }
  try {
    const saved = await createWorkLog({
      taskTitle,
      project,
      hours,
      category,
      status,
      commit,
      date: date || new Date().toISOString().slice(0, 10),
      submittedBy,
    });
    return res.status(201).json(saved);
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || "Failed to save the work log entry." });
  }
});

app.delete("/api/worklogs/:id", async (req, res) => {
  if (!isGraphAppConfigured()) return res.status(501).json(NOT_CONFIGURED_ERROR);
  try {
    await removeWorkLog(req.params.id);
    return res.status(204).end();
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || "Failed to delete the work log entry." });
  }
});

app.listen(PORT, () => {
  console.log(`Bench auth server listening on http://localhost:${PORT}`);
});
