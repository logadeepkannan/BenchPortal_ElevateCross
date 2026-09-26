// App-only equivalent of the Work Log section of src/lib/sharePointService.js,
// used for the manual (email) sign-in path, which has no delegated Graph
// token of its own — only this server's client-credential grant. Field
// mapping (ProjectName as a Lookup resolved by title, EntryDate as an ISO
// date, etc.) is kept identical to the delegated client so entries written
// by either sign-in path look the same in BenchWorkLogs.
const {
  getAppToken,
  getSharePointAppToken,
  spRestFetch,
  resolveSiteId,
  resolveListId,
  getListItems,
  createListItem,
  updateListItem,
  deleteListItem,
  getColumnNameMap,
} = require("./graphAppClient");

const LIST_NAMES = { workLogs: "BenchWorkLogs", projects: "BenchProjects" };

let siteIdCache = null;
async function getSiteId(token) {
  if (!siteIdCache) siteIdCache = await resolveSiteId(token);
  return siteIdCache;
}

const listIdCache = new Map();
async function getListId(token, siteId, displayName) {
  if (!listIdCache.has(displayName)) {
    listIdCache.set(displayName, await resolveListId(token, siteId, displayName));
  }
  return listIdCache.get(displayName);
}

let projectLookupCache = null; // { byTitle: Map<title, id>, byId: Map<id, title> }
async function getProjectLookup(token, siteId, { forceRefresh = false } = {}) {
  if (projectLookupCache && !forceRefresh) return projectLookupCache;
  const listId = await getListId(token, siteId, LIST_NAMES.projects);
  const items = await getListItems(token, siteId, listId);
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

// ProjectManager, TechnicalLead and TeamMembers are real Person or Group
// columns on BenchProjects. Graph reads/writes them the same way as a Lookup
// column ("<Field>LookupId"), but that id belongs to the site's own User
// Information List, and Graph has no operation that resolves a name or email
// to an id there — so it goes through SharePoint REST's siteusers endpoint
// instead (read-only) and is cached; only people who've opened the SharePoint
// site at least once show up here.
let siteUsersCache = null; // { byTitle: Map<title, id>, byId: Map<id, title> }
async function getSiteUsersCache({ forceRefresh = false } = {}) {
  if (siteUsersCache && !forceRefresh) return siteUsersCache;
  const spToken = await getSharePointAppToken();
  const result = await spRestFetch(spToken, `/_api/web/siteusers?$filter=PrincipalType eq 1`);
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

async function fetchSiteUsers() {
  const { byTitle } = await getSiteUsersCache({ forceRefresh: true });
  return Array.from(byTitle.keys());
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
// through getColumnNameMap (same fix already used for BenchAdministration's
// manual-login lookup above) avoids guessing it.
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
async function getProjectFieldNames(token, siteId, listId) {
  if (projectFieldNamesCache) return projectFieldNamesCache;
  const columnMap = await getColumnNameMap(token, siteId, listId);
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

async function fetchProjects() {
  const token = await getAppToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(token, siteId, LIST_NAMES.projects);
  const [items, { byId: peopleById }, names] = await Promise.all([
    getListItems(token, siteId, listId),
    getSiteUsersCache({ forceRefresh: true }),
    getProjectFieldNames(token, siteId, listId),
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

async function createProject(project) {
  const token = await getAppToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(token, siteId, LIST_NAMES.projects);
  const [{ byTitle: peopleByTitle }, names] = await Promise.all([getSiteUsersCache(), getProjectFieldNames(token, siteId, listId)]);
  const fields = projectToFields(project, peopleByTitle, names);
  const item = await createListItem(token, siteId, listId, fields);
  projectLookupCache = null;
  return { ...project, id: item.id };
}

async function updateProject(id, project) {
  const token = await getAppToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(token, siteId, LIST_NAMES.projects);
  const [{ byTitle: peopleByTitle }, names] = await Promise.all([getSiteUsersCache(), getProjectFieldNames(token, siteId, listId)]);
  const fields = projectToFields(project, peopleByTitle, names);
  await updateListItem(token, siteId, listId, id, fields);
  projectLookupCache = null;
  return { ...project, id };
}

async function removeProject(id) {
  const token = await getAppToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(token, siteId, LIST_NAMES.projects);
  await deleteListItem(token, siteId, listId, id);
  projectLookupCache = null;
}

// JiraTicketLink was added after BenchWorkLogs' other columns, which were
// typed without spaces so their internal names happen to match ("HoursSpent",
// "GitCommitSignature", …). "Jira Ticket Link" has spaces, so — same issue
// already fixed for BenchProjects — its internal name won't just be
// "JiraTicketLink"; resolve it through getColumnNameMap instead of guessing.
let workLogFieldNamesCache = null;
async function getWorkLogFieldNames(token, siteId, listId) {
  if (workLogFieldNamesCache) return workLogFieldNamesCache;
  const columnMap = await getColumnNameMap(token, siteId, listId);
  workLogFieldNamesCache = { jiraLink: columnMap.get("Jira Ticket Link") || "JiraTicketLink" };
  return workLogFieldNamesCache;
}

function workLogFromItem(item, byId, names) {
  const f = item.fields || {};
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

async function fetchWorkLogs() {
  const token = await getAppToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(token, siteId, LIST_NAMES.workLogs);
  const [items, { byId }, names] = await Promise.all([
    getListItems(token, siteId, listId),
    getProjectLookup(token, siteId, { forceRefresh: true }),
    getWorkLogFieldNames(token, siteId, listId),
  ]);
  return items.map((item) => workLogFromItem(item, byId, names));
}

async function createWorkLog(log) {
  const token = await getAppToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(token, siteId, LIST_NAMES.workLogs);
  const [{ byTitle, byId }, names] = await Promise.all([
    getProjectLookup(token, siteId),
    getWorkLogFieldNames(token, siteId, listId),
  ]);
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
  const item = await createListItem(token, siteId, listId, fields);
  return workLogFromItem(item, byId, names);
}

async function removeWorkLog(itemId) {
  const token = await getAppToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(token, siteId, LIST_NAMES.workLogs);
  await deleteListItem(token, siteId, listId, itemId);
}

module.exports = { fetchProjects, createProject, updateProject, removeProject, fetchSiteUsers, fetchWorkLogs, createWorkLog, removeWorkLog };
