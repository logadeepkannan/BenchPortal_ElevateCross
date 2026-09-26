import { getGraphAccessToken, getSharePointRestToken, isMsalConfigured } from "./msal";

const SP_HOSTNAME = import.meta.env.VITE_SP_HOSTNAME || "";
const SP_SITE_PATH = import.meta.env.VITE_SP_SITE_PATH || "";

export const LIST_NAMES = {
  workLogs: "BenchWorkLogs",
  issues: "BenchIssueTracker",
  projects: "BenchProjects",
  administration: "BenchAdministration",
};

export const ACCESS_LEVELS = ["Admin", "Manager", "User"];
const DEFAULT_ACCESS_LEVEL = "User";

const COLUMN_LABELS = { todo: "To Do", inprogress: "In Progress", critical: "Critical / At Risk" };
const COLUMN_KEYS = Object.fromEntries(Object.entries(COLUMN_LABELS).map(([k, v]) => [v, k]));

export function isSharePointConfigured() {
  return isMsalConfigured() && Boolean(SP_HOSTNAME && SP_SITE_PATH);
}

/* ----------------------------- Graph plumbing ---------------------------- */

async function graphFetch(path, options = {}) {
  const token = await getGraphAccessToken();
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

let siteIdCache = null;
async function resolveSiteId() {
  if (siteIdCache) return siteIdCache;
  if (!SP_HOSTNAME || !SP_SITE_PATH) {
    throw new Error("SharePoint site is not configured. Set VITE_SP_HOSTNAME and VITE_SP_SITE_PATH in .env.");
  }
  const path = SP_SITE_PATH.startsWith("/") ? SP_SITE_PATH : `/${SP_SITE_PATH}`;
  const site = await graphFetch(`/sites/${SP_HOSTNAME}:${path}`);
  siteIdCache = site.id;
  return siteIdCache;
}

const listIdCache = new Map();
async function resolveListId(displayName) {
  if (listIdCache.has(displayName)) return listIdCache.get(displayName);
  const siteId = await resolveSiteId();
  const result = await graphFetch(`/sites/${siteId}/lists?$filter=displayName eq '${displayName}'`);
  const list = result.value && result.value[0];
  if (!list) {
    throw new Error(`SharePoint list "${displayName}" was not found on the configured site. Create it using the schema in Architecture (Admin).`);
  }
  listIdCache.set(displayName, list.id);
  return list.id;
}

async function getSiteAndList(displayName) {
  const siteId = await resolveSiteId();
  const listId = await resolveListId(displayName);
  return { siteId, listId };
}

// Lists created through the modern Microsoft Lists app frequently store
// columns under generic internal names (e.g. "field_2") that don't match the
// display name you typed and see in the UI ("Access_Level") — Graph's fields
// resource is keyed by that internal name, not the display name. Any code
// that reads/writes a column by assuming they match will silently fail (the
// field just comes back undefined) rather than error, so this resolves the
// real mapping from the list's own column definitions instead of guessing.
const columnMapCache = new Map(); // displayName (list) -> Map<column displayName, internal name>
async function getColumnNameMap(listDisplayName) {
  if (columnMapCache.has(listDisplayName)) return columnMapCache.get(listDisplayName);
  const { siteId, listId } = await getSiteAndList(listDisplayName);
  const result = await graphFetch(`/sites/${siteId}/lists/${listId}/columns`);
  const map = new Map();
  (result.value || []).forEach((col) => {
    if (col.displayName) map.set(col.displayName, col.name);
  });
  columnMapCache.set(listDisplayName, map);
  return map;
}

async function getListItems(displayName) {
  const { siteId, listId } = await getSiteAndList(displayName);
  const result = await graphFetch(`/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500&$orderby=fields/Created desc`);
  return result.value || [];
}

async function createListItem(displayName, fields) {
  const { siteId, listId } = await getSiteAndList(displayName);
  return graphFetch(`/sites/${siteId}/lists/${listId}/items`, {
    method: "POST",
    body: JSON.stringify({ fields }),
  });
}

async function updateListItemFields(displayName, itemId, fields) {
  const { siteId, listId } = await getSiteAndList(displayName);
  return graphFetch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, {
    method: "PATCH",
    body: JSON.stringify(fields),
  });
}

async function deleteListItem(displayName, itemId) {
  const { siteId, listId } = await getSiteAndList(displayName);
  await graphFetch(`/sites/${siteId}/lists/${listId}/items/${itemId}`, { method: "DELETE" });
}

/* ------------------------- SharePoint REST plumbing ------------------------ */
/* Read-only: used only to resolve Person/Group column ids (see below). All
 * writes still go through Graph, including for the Person column itself. */

async function spRestFetch(path) {
  const token = await getSharePointRestToken(SP_HOSTNAME);
  const sitePath = SP_SITE_PATH.startsWith("/") ? SP_SITE_PATH : `/${SP_SITE_PATH}`;
  const res = await fetch(`https://${SP_HOSTNAME}${sitePath}${path}`, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=nometadata",
    },
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`SharePoint REST ${res.status} ${res.statusText}: ${body.slice(0, 300)}`);
  }
  return res.json();
}

/* --------------------------- Field mapping: Work Logs --------------------------- */
/*
 * ProjectName is a real Lookup column pointing at the BenchProjects list.
 * Graph reads/writes Lookup columns as "{InternalName}LookupId" — the same
 * item id Graph itself returns for items in the target list — so resolving it
 * is just a title->id map built from BenchProjects, no SharePoint REST needed.
 * GitCommitSignature stays plain text. SubmittedBy is a plain Single line of
 * text column holding the submitter's email — written explicitly rather than
 * relying on Author/Created By, since manual (email) sign-in writes through
 * this app's own Graph app-only credential, so Created By would always show
 * that service identity rather than the actual person.
 */

let projectLookupCache = null; // { byTitle: Map<title, id>, byId: Map<id, title> }

async function getProjectLookupCache({ forceRefresh = false } = {}) {
  if (projectLookupCache && !forceRefresh) return projectLookupCache;
  const items = await getListItems(LIST_NAMES.projects);
  const byTitle = new Map();
  const byId = new Map();
  items.forEach((item) => {
    const title = item.fields?.Title;
    if (!title) return;
    byTitle.set(title, item.id);
    byId.set(String(item.id), title);
  });
  projectLookupCache = { byTitle, byId };
  return projectLookupCache;
}

function resolvePersonId(byTitle, name) {
  if (!name) return null;
  const id = byTitle.get(name);
  if (!id) {
    throw new Error(`"${name}" isn't a recognized site user yet. Pick someone from the list, or have them open the SharePoint site once first.`);
  }
  return id;
}

// BenchProjects' columns were created by typing a display name with spaces
// ("Project Code", "Client / Company", …) straight into the SharePoint UI,
// which auto-generates an internal name that does NOT match — e.g. "Project
// Code" becomes "Project_x0020_Code", not "ProjectCode". Graph's item.fields
// is keyed by that real internal name, so writing/reading with the assumed
// PascalCase name silently targets a field that doesn't exist. Resolving
// through getColumnNameMap (same fix already used for BenchAdministration)
// avoids guessing it.
const PROJECT_FIELD_DISPLAY_NAMES = {
  code: "Project Code",
  description: "Description",
  status: "Status",
  health: "Project Health",
  category: "Category",
  projectManager: "Project Manager",
  technicalLead: "Technical Lead",
  teamMembers: "Team Members",
  client: "Client / Company",
  zone: "Zone",
  startDate: "Start Date",
  targetEndDate: "Target End Date",
  actualEndDate: "Actual End Date",
  estimatedHours: "Estimated Hours",
  repositoryUrl: "Repository URL",
  documentationLink: "Documentation Link",
};

let projectFieldNamesCache = null;
async function getProjectFieldNames() {
  if (projectFieldNamesCache) return projectFieldNamesCache;
  const columnMap = await getColumnNameMap(LIST_NAMES.projects);
  projectFieldNamesCache = Object.fromEntries(
    Object.entries(PROJECT_FIELD_DISPLAY_NAMES).map(([key, display]) => [key, columnMap.get(display) || display])
  );
  return projectFieldNamesCache;
}

function projectToFields(project, byTitle, names) {
  const fields = {
    Title: project.title,
    [names.code]: project.code || "",
    [names.description]: project.description || "",
    [names.status]: project.status || "",
    [names.health]: project.health || "",
    [names.category]: project.category || "",
    [names.client]: project.client || "",
    [names.zone]: project.zone || "",
    [names.estimatedHours]: Number(project.estimatedHours) || 0,
  };
  if (project.projectManager) fields[`${names.projectManager}LookupId`] = resolvePersonId(byTitle, project.projectManager);
  if (project.technicalLead) fields[`${names.technicalLead}LookupId`] = resolvePersonId(byTitle, project.technicalLead);
  if (project.teamMembers && project.teamMembers.length) {
    fields[`${names.teamMembers}LookupId`] = project.teamMembers.map((name) => resolvePersonId(byTitle, name));
  }
  if (project.startDate) fields[names.startDate] = `${project.startDate}T00:00:00Z`;
  if (project.targetEndDate) fields[names.targetEndDate] = `${project.targetEndDate}T00:00:00Z`;
  if (project.actualEndDate) fields[names.actualEndDate] = `${project.actualEndDate}T00:00:00Z`;
  if (project.repositoryUrl) fields[names.repositoryUrl] = { Url: project.repositoryUrl };
  if (project.documentationLink) fields[names.documentationLink] = { Url: project.documentationLink };
  return fields;
}

function projectFromItem(item, peopleById, names) {
  const f = item.fields || {};
  const personName = (lookupId) => (lookupId != null ? peopleById.get(String(lookupId)) || "" : "");
  const teamIds = Array.isArray(f[`${names.teamMembers}LookupId`]) ? f[`${names.teamMembers}LookupId`] : [];
  return {
    id: item.id,
    title: f.Title || "",
    code: f[names.code] || "",
    description: f[names.description] || "",
    status: f[names.status] || "",
    health: f[names.health] || "",
    category: f[names.category] || "",
    projectManager: personName(f[`${names.projectManager}LookupId`]),
    technicalLead: personName(f[`${names.technicalLead}LookupId`]),
    teamMembers: teamIds.map((id) => peopleById.get(String(id))).filter(Boolean),
    client: f[names.client] || "",
    zone: f[names.zone] || "",
    startDate: (f[names.startDate] || "").slice(0, 10),
    targetEndDate: (f[names.targetEndDate] || "").slice(0, 10),
    actualEndDate: (f[names.actualEndDate] || "").slice(0, 10),
    estimatedHours: Number(f[names.estimatedHours]) || 0,
    repositoryUrl: f[names.repositoryUrl]?.Url || "",
    documentationLink: f[names.documentationLink]?.Url || "",
  };
}

export async function fetchProjects() {
  const [items, { byId: peopleById }, names] = await Promise.all([
    getListItems(LIST_NAMES.projects),
    getSiteUsersCache({ forceRefresh: true }),
    getProjectFieldNames(),
  ]);
  const byTitle = new Map();
  const byId = new Map();
  items.forEach((item) => {
    const title = item.fields?.Title;
    if (title) {
      byTitle.set(title, item.id);
      byId.set(String(item.id), title);
    }
  });
  projectLookupCache = { byTitle, byId };
  return items.map((item) => projectFromItem(item, peopleById, names));
}

export async function createProject(project) {
  const [{ byTitle: peopleByTitle }, names] = await Promise.all([getSiteUsersCache(), getProjectFieldNames()]);
  const fields = projectToFields(project, peopleByTitle, names);
  const item = await createListItem(LIST_NAMES.projects, fields);
  projectLookupCache = null;
  return { ...project, id: item.id };
}

export async function updateProject(id, project) {
  const [{ byTitle: peopleByTitle }, names] = await Promise.all([getSiteUsersCache(), getProjectFieldNames()]);
  const fields = projectToFields(project, peopleByTitle, names);
  await updateListItemFields(LIST_NAMES.projects, id, fields);
  projectLookupCache = null;
  return { ...project, id };
}

export async function removeProject(id) {
  await deleteListItem(LIST_NAMES.projects, id);
  projectLookupCache = null;
}

// JiraTicketLink was added after BenchWorkLogs' other columns, which were
// typed without spaces so their internal names happen to match ("HoursSpent",
// "GitCommitSignature", …). "Jira Ticket Link" has spaces, so — same issue
// already fixed for BenchProjects — its internal name won't just be
// "JiraTicketLink"; resolve it through getColumnNameMap instead of guessing.
let workLogFieldNamesCache = null;
async function getWorkLogFieldNames() {
  if (workLogFieldNamesCache) return workLogFieldNamesCache;
  const columnMap = await getColumnNameMap(LIST_NAMES.workLogs);
  workLogFieldNamesCache = { jiraLink: columnMap.get("Jira Ticket Link") || "JiraTicketLink" };
  return workLogFieldNamesCache;
}

async function workLogToFields(log) {
  const [{ byTitle }, names] = await Promise.all([getProjectLookupCache(), getWorkLogFieldNames()]);
  const projectId = byTitle.get(log.project);
  if (!projectId) {
    throw new Error(`Project "${log.project}" was not found in the BenchProjects list. Add it there first, then try again.`);
  }
  const fields = {
    Title: log.taskTitle,
    ProjectNameLookupId: projectId,
    HoursSpent: Number(log.hours) || 0,
    Category: log.category,
    Status: log.status,
    GitCommitSignature: log.commit && log.commit !== "-" ? log.commit : "",
    EntryDate: `${log.date}T00:00:00Z`,
    SubmittedBy: log.submittedBy || "",
  };
  if (log.jiraLink) fields[names.jiraLink] = { Url: log.jiraLink };
  return fields;
}

async function workLogFromItem(item) {
  const f = item.fields || {};
  const [{ byId }, names] = await Promise.all([getProjectLookupCache(), getWorkLogFieldNames()]);
  return {
    id: item.id,
    date: (f.EntryDate || "").slice(0, 10) || new Date().toISOString().slice(0, 10),
    taskTitle: f.Title || "(untitled)",
    project: byId.get(String(f.ProjectNameLookupId)) || "(unknown project)",
    hours: Number(f.HoursSpent) || 0,
    category: f.Category || "Development",
    status: f.Status || "In Progress",
    commit: f.GitCommitSignature || "-",
    submittedBy: f.SubmittedBy || "",
    jiraLink: f[names.jiraLink]?.Url || "",
  };
}

export async function fetchWorkLogs() {
  const items = await getListItems(LIST_NAMES.workLogs);
  await getProjectLookupCache({ forceRefresh: true });
  return Promise.all(items.map(workLogFromItem));
}

export async function createWorkLog(log) {
  const fields = await workLogToFields(log);
  const item = await createListItem(LIST_NAMES.workLogs, fields);
  return workLogFromItem(item);
}

export async function removeWorkLog(itemId) {
  await deleteListItem(LIST_NAMES.workLogs, itemId);
}

/* --------------------------- Shared: site users (Person columns) --------------------------- */
/*
 * AssignedTo (Issue Tracker) and ProjectManager / TechnicalLead / TeamMembers
 * (Projects) are all real Person or Group columns. Graph writes/reads them
 * the same way as a Lookup column — "<Field>LookupId" — but that id belongs
 * to the site's own User Information List, and Graph has no operation that
 * resolves a name or email to an id in that list. So the id lookup goes
 * through SharePoint REST's siteusers endpoint instead (read-only) and is
 * cached; only people already present in that list can be picked — there's
 * no _api/web/ensureuser fallback here, so if someone new needs to be
 * pickable, have them open the SharePoint site once first.
 */

let siteUsersCache = null; // { byTitle: Map<title, id>, byId: Map<id, title> }

async function getSiteUsersCache({ forceRefresh = false } = {}) {
  if (siteUsersCache && !forceRefresh) return siteUsersCache;
  const result = await spRestFetch(`/_api/web/siteusers?$filter=PrincipalType eq 1`);
  const byTitle = new Map();
  const byId = new Map();
  (result.value || []).forEach((u) => {
    if (!u.Title || !u.Email) return;
    byTitle.set(u.Title, u.Id);
    byId.set(String(u.Id), u.Title);
  });
  siteUsersCache = { byTitle, byId };
  return siteUsersCache;
}

export async function fetchAssignees() {
  const { byTitle } = await getSiteUsersCache({ forceRefresh: true });
  return Array.from(byTitle.keys());
}

/* --------------------------- Field mapping: Issue Tracker --------------------------- */
/* Tags stays a plain "; "-joined string rather than real Managed Metadata. */

async function issueToFields(issue) {
  const { byTitle } = await getSiteUsersCache();
  const assigneeId = byTitle.get(issue.assignee);
  if (!assigneeId) {
    throw new Error(`"${issue.assignee}" isn't a recognized site user yet. Pick someone from the list, or have them open the SharePoint site once first.`);
  }
  return {
    Title: issue.title,
    Description: issue.description || "",
    AssignedToLookupId: assigneeId,
    Priority: issue.priority,
    DueDate: `${issue.dueDate}T00:00:00Z`,
    BoardColumn: COLUMN_LABELS[issue.column] || COLUMN_LABELS.todo,
    Tags: (issue.tags || []).join("; "),
  };
}

async function issueFromItem(item) {
  const f = item.fields || {};
  const { byId } = await getSiteUsersCache();
  return {
    id: item.id,
    title: f.Title || "(untitled)",
    description: f.Description || "",
    assignee: byId.get(String(f.AssignedToLookupId)) || "Unassigned",
    priority: f.Priority || "Medium",
    dueDate: (f.DueDate || "").slice(0, 10),
    tags: f.Tags
      ? String(f.Tags)
          .split(";")
          .map((t) => t.trim())
          .filter(Boolean)
      : [],
    column: COLUMN_KEYS[f.BoardColumn] || "todo",
  };
}

export async function fetchIssues() {
  const items = await getListItems(LIST_NAMES.issues);
  await getSiteUsersCache({ forceRefresh: true });
  return Promise.all(items.map(issueFromItem));
}

export async function createIssue(issue) {
  const fields = await issueToFields(issue);
  const item = await createListItem(LIST_NAMES.issues, fields);
  return issueFromItem(item);
}

export async function moveIssue(itemId, column) {
  await updateListItemFields(LIST_NAMES.issues, itemId, { BoardColumn: COLUMN_LABELS[column] || COLUMN_LABELS.todo });
}

/* --------------------------- Access control: BenchAdministration --------------------------- */
/*
 * A simple role list — UserName, UserEmail, Access_Level (Admin / Manager /
 * User) — keyed by email. Not a security boundary on its own (it only hides
 * the Architecture (Admin) nav item in the UI; it doesn't restrict what the
 * signed-in Graph token itself can read/write), but it's the right layer for
 * "who should see the admin screens" in a reference app like this one.
 */

export async function fetchAccessLevel(email) {
  const normalized = (email || "").trim().toLowerCase();
  if (!normalized) return DEFAULT_ACCESS_LEVEL;
  const [columnMap, items] = await Promise.all([getColumnNameMap(LIST_NAMES.administration), getListItems(LIST_NAMES.administration)]);
  const emailField = columnMap.get("UserEmail") || "UserEmail";
  const levelField = columnMap.get("Access_Level") || "Access_Level";
  const match = items.find((item) => String(item.fields?.[emailField] || "").trim().toLowerCase() === normalized);
  const level = match?.fields?.[levelField];
  return ACCESS_LEVELS.includes(level) ? level : DEFAULT_ACCESS_LEVEL;
}

// Called on every Microsoft sign-in. Looks the person up in BenchAdministration
// same as fetchAccessLevel, but if they've never signed in before, provisions
// a new row for them (name, email, Access_Level = "User") instead of just
// defaulting silently — so a first-time sign-in shows up in the list ready
// for an admin to promote, rather than needing to be added by hand up front.
export async function ensureAdministrationEntry({ name, email }) {
  const normalized = (email || "").trim().toLowerCase();
  if (!normalized) return DEFAULT_ACCESS_LEVEL;

  const [columnMap, items] = await Promise.all([getColumnNameMap(LIST_NAMES.administration), getListItems(LIST_NAMES.administration)]);
  const nameField = columnMap.get("UserName") || "UserName";
  const emailField = columnMap.get("UserEmail") || "UserEmail";
  const levelField = columnMap.get("Access_Level") || "Access_Level";

  const existing = items.find((item) => String(item.fields?.[emailField] || "").trim().toLowerCase() === normalized);
  if (existing) {
    const level = existing.fields?.[levelField];
    return ACCESS_LEVELS.includes(level) ? level : DEFAULT_ACCESS_LEVEL;
  }

  await createListItem(LIST_NAMES.administration, {
    [nameField]: name || email,
    [emailField]: email,
    [levelField]: DEFAULT_ACCESS_LEVEL,
  });
  return DEFAULT_ACCESS_LEVEL;
}

// MSAL's account.username is the sign-in UPN, which isn't always the same
// address as the user's actual mailbox — tenants commonly issue UPNs on a
// different domain than the primary email. If BenchAdministration was
// populated with real mailbox addresses, comparing against the UPN would
// never match. /me's "mail" property is the actual mailbox address; fall
// back to the UPN only for the (rarer) accounts that don't have one set.
export async function fetchMyProfile() {
  const me = await graphFetch(`/me?$select=displayName,mail,userPrincipalName`);
  return {
    name: me.displayName || "",
    email: me.mail || me.userPrincipalName || "",
  };
}
