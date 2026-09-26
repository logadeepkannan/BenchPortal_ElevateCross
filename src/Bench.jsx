import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import { createPortal } from "react-dom";
import {
  LayoutDashboard,
  ClipboardList,
  BookOpen,
  AlertTriangle,
  GraduationCap,
  Code2,
  KanbanSquare,
  Settings2,
  Sun,
  Moon,
  Menu,
  X,
  Search,
  Plus,
  Download,
  Copy,
  Check,
  ChevronDown,
  ChevronRight,
  ExternalLink,
  Clock,
  TrendingUp,
  Award,
  CheckCircle2,
  Bell,
  ShieldAlert,
  ShieldCheck,
  Sparkles,
  Send,
  Loader2,
  Trash2,
  Calendar,
  GitCommit,
  Tag,
  Users,
  Server,
  Database,
  FolderTree,
  Workflow,
  BarChart3,
  Zap,
  BookMarked,
  FileCode2,
  Layers,
  MoreVertical,
  CircleUserRound,
  Rocket,
  PanelLeftClose,
  PanelLeftOpen,
  Flame,
  Cloud,
  LogOut,
  Mail,
  ArrowRight,
  RefreshCw,
  CloudOff,
  AlertCircle,
  Briefcase,
  Info,
  Link2,
  Pencil,
} from "lucide-react";
import { signInWithMicrosoft, completeMicrosoftSignIn, signOutMicrosoft, isMsalConfigured } from "./lib/msal";
import { verifyManualLogin } from "./lib/manualAuth";
import {
  isSharePointConfigured,
  fetchWorkLogs,
  createWorkLog,
  removeWorkLog,
  fetchIssues,
  createIssue,
  moveIssue,
  fetchProjects,
  createProject,
  updateProject,
  removeProject,
  fetchAssignees,
  ensureAdministrationEntry,
  fetchMyProfile,
} from "./lib/sharePointService";
import {
  isManualWorkLogAvailable,
  fetchManualProjects,
  createManualProject,
  updateManualProject,
  removeManualProject,
  fetchManualSiteUsers,
  fetchManualWorkLogs,
  createManualWorkLog,
  removeManualWorkLog,
} from "./lib/manualWorkLogService";

/* ----------------------------------------------------------------------- */
/* Utilities                                                                */
/* ----------------------------------------------------------------------- */

function cx(...parts) {
  return parts.filter(Boolean).join(" ");
}

function isoDaysAgo(n) {
  const d = new Date();
  d.setDate(d.getDate() - n);
  return d.toISOString().slice(0, 10);
}

function formatDate(iso) {
  const d = new Date(iso + "T00:00:00");
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}

function formatDateShort(iso) {
  const d = new Date(iso + "T00:00:00");
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric" });
}

function weekdayLabel(iso) {
  const d = new Date(iso + "T00:00:00");
  return d.toLocaleDateString("en-US", { weekday: "short" });
}

function downloadCSV(rows, filename) {
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  const escape = (val) => {
    const s = String(val ?? "");
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const csv = [headers.join(","), ...rows.map((r) => headers.map((h) => escape(r[h])).join(","))].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function useCopyToClipboard() {
  const [copiedKey, setCopiedKey] = useState(null);
  const copy = useCallback((text, key) => {
    const fallback = () => {
      const ta = document.createElement("textarea");
      ta.value = text;
      ta.style.position = "fixed";
      ta.style.opacity = "0";
      document.body.appendChild(ta);
      ta.select();
      try {
        document.execCommand("copy");
      } catch (e) {
        /* no-op */
      }
      document.body.removeChild(ta);
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).catch(fallback);
    } else {
      fallback();
    }
    setCopiedKey(key);
    setTimeout(() => setCopiedKey((k) => (k === key ? null : k)), 1800);
  }, []);
  return [copiedKey, copy];
}

/* ----------------------------------------------------------------------- */
/* Seed data                                                                */
/* ----------------------------------------------------------------------- */

const PROJECTS = ["Contoso Intranet Hub", "Fabrikam SPFx Suite", "Northwind Identity Gateway", "Adventure Works Analytics"];
const ASSIGNEES = ["Priya N.", "Dana W.", "Marcus T.", "Alex Rivera"];
const PROJECT_STATUSES = ["Planning", "Active", "On Hold", "Completed", "Cancelled"];
const PROJECT_HEALTHS = ["Green", "Yellow", "Red"];
const PROJECT_CATEGORIES = ["Client", "Internal", "R&D"];
const EMPTY_PROJECT_FORM = {
  title: "",
  code: "",
  description: "",
  status: PROJECT_STATUSES[1],
  health: PROJECT_HEALTHS[0],
  category: PROJECT_CATEGORIES[0],
  projectManager: "",
  technicalLead: "",
  teamMembers: [],
  client: "",
  zone: "",
  startDate: isoDaysAgo(0),
  targetEndDate: isoDaysAgo(-90),
  actualEndDate: "",
  estimatedHours: "",
  repositoryUrl: "",
  documentationLink: "",
};
function projectToForm(p) {
  return {
    title: p.title || "",
    code: p.code || "",
    description: p.description || "",
    status: p.status || PROJECT_STATUSES[1],
    health: p.health || PROJECT_HEALTHS[0],
    category: p.category || PROJECT_CATEGORIES[0],
    projectManager: p.projectManager || "",
    technicalLead: p.technicalLead || "",
    teamMembers: p.teamMembers || [],
    client: p.client || "",
    zone: p.zone || "",
    startDate: p.startDate || "",
    targetEndDate: p.targetEndDate || "",
    actualEndDate: p.actualEndDate || "",
    estimatedHours: p.estimatedHours || "",
    repositoryUrl: p.repositoryUrl || "",
    documentationLink: p.documentationLink || "",
  };
}
const SEED_PROJECTS = [
  {
    id: "local-0",
    title: "Contoso Intranet Hub",
    code: "CTH-001",
    description: "Modern SharePoint intranet home site with Graph-powered personalization.",
    status: "Active",
    health: "Green",
    category: "Client",
    projectManager: "Priya N.",
    technicalLead: "Marcus T.",
    teamMembers: ["Marcus T.", "Dana W."],
    client: "Contoso Ltd.",
    zone: "NA-East",
    startDate: isoDaysAgo(120),
    targetEndDate: isoDaysAgo(-30),
    actualEndDate: "",
    estimatedHours: 480,
    repositoryUrl: "https://github.com/contoso/intranet-hub",
    documentationLink: "https://contoso.sharepoint.com/sites/IntranetHub/docs",
  },
  {
    id: "local-1",
    title: "Fabrikam SPFx Suite",
    code: "FSS-002",
    description: "Collection of SPFx web parts and extensions for Fabrikam's field operations portal.",
    status: "Active",
    health: "Yellow",
    category: "Client",
    projectManager: "Dana W.",
    technicalLead: "Alex Rivera",
    teamMembers: ["Alex Rivera", "Priya N."],
    client: "Fabrikam Inc.",
    zone: "NA-West",
    startDate: isoDaysAgo(200),
    targetEndDate: isoDaysAgo(-10),
    actualEndDate: "",
    estimatedHours: 620,
    repositoryUrl: "https://github.com/fabrikam/spfx-suite",
    documentationLink: "",
  },
  {
    id: "local-2",
    title: "Northwind Identity Gateway",
    code: "NIG-003",
    description: "Entra ID conditional access and Graph subscription webhook gateway.",
    status: "On Hold",
    health: "Red",
    category: "Internal",
    projectManager: "Marcus T.",
    technicalLead: "Priya N.",
    teamMembers: ["Priya N."],
    client: "",
    zone: "EMEA",
    startDate: isoDaysAgo(260),
    targetEndDate: isoDaysAgo(20),
    actualEndDate: "",
    estimatedHours: 340,
    repositoryUrl: "",
    documentationLink: "https://contoso.sharepoint.com/sites/NorthwindGateway/docs",
  },
  {
    id: "local-3",
    title: "Adventure Works Analytics",
    code: "AWA-004",
    description: "Power Platform analytics workspace for delivery leadership reporting.",
    status: "Completed",
    health: "Green",
    category: "R&D",
    projectManager: "Priya N.",
    technicalLead: "Dana W.",
    teamMembers: ["Dana W.", "Marcus T.", "Alex Rivera"],
    client: "",
    zone: "NA-East",
    startDate: isoDaysAgo(400),
    targetEndDate: isoDaysAgo(60),
    actualEndDate: isoDaysAgo(45),
    estimatedHours: 210,
    repositoryUrl: "https://github.com/adventure-works/analytics",
    documentationLink: "",
  },
];
const WORK_CATEGORIES = ["Development", "Code Review", "Meeting", "Documentation", "Testing", "Deployment", "Research"];
const WORK_STATUSES = ["Completed", "In Progress", "Blocked"];

const SEED_WORK_LOGS = [
  { id: 1, date: isoDaysAgo(0), taskTitle: "Implement Microsoft Graph batch requests for profile sync", project: "Contoso Intranet Hub", hours: 4.5, category: "Development", status: "Completed", commit: "a3f9e21", jiraLink: "https://bench.atlassian.net/browse/CTH-142" },
  { id: 2, date: isoDaysAgo(0), taskTitle: "Code review: SPFx property pane refactor", project: "Fabrikam SPFx Suite", hours: 1.5, category: "Code Review", status: "Completed", commit: "-", jiraLink: "" },
  { id: 3, date: isoDaysAgo(1), taskTitle: "Sprint planning & backlog grooming", project: "Contoso Intranet Hub", hours: 1, category: "Meeting", status: "Completed", commit: "-", jiraLink: "" },
  { id: 4, date: isoDaysAgo(1), taskTitle: "Fix Entra ID conditional access token refresh bug", project: "Northwind Identity Gateway", hours: 5, category: "Development", status: "Completed", commit: "7c1d4b8", jiraLink: "https://bench.atlassian.net/browse/NIG-58" },
  { id: 5, date: isoDaysAgo(2), taskTitle: "Write Power Automate flow documentation", project: "Contoso Intranet Hub", hours: 2, category: "Documentation", status: "Completed", commit: "-", jiraLink: "" },
  { id: 6, date: isoDaysAgo(3), taskTitle: "Investigate list view threshold throttling", project: "Fabrikam SPFx Suite", hours: 3, category: "Research", status: "In Progress", commit: "-", jiraLink: "https://bench.atlassian.net/browse/FSS-91" },
  { id: 7, date: isoDaysAgo(4), taskTitle: "UAT regression testing for release 4.2", project: "Northwind Identity Gateway", hours: 6, category: "Testing", status: "Completed", commit: "-", jiraLink: "" },
  { id: 8, date: isoDaysAgo(5), taskTitle: "Deploy hotfix to production SPFx package", project: "Fabrikam SPFx Suite", hours: 1.5, category: "Deployment", status: "Blocked", commit: "f02aa19", jiraLink: "" },
];

const KB_ARTICLES = [
  {
    id: "kb-1",
    title: "Resolving CORS errors when calling Microsoft Graph from SPFx",
    category: "Integrations",
    tags: ["Graph API", "SPFx", "CORS"],
    problem: "Custom SPFx web part throws a CORS / 403 error when calling Microsoft Graph endpoints directly via fetch().",
    rootCause: "SPFx web parts must acquire tokens through the AadHttpClient / MSGraphClientV3 factory so the correct audience and API permissions from the app registration are used. Raw fetch() calls bypass SPO's token proxy and are blocked by the browser's CORS policy.",
    solution: "Use the AadHttpClientFactory or MSGraphClientFactory injected via the web part context instead of raw fetch. Grant the Graph delegated permission in the SharePoint API access admin center, then approve it before the client is usable.",
    language: "typescript",
    code:
      "import { MSGraphClientV3 } from \"@microsoft/sp-http\";\n\nexport async function getMyRecentFiles(context) {\n  const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient(\"3\");\n  const response = await client\n    .api(\"/me/drive/recent\")\n    .version(\"v1.0\")\n    .top(10)\n    .get();\n  return response.value;\n}",
    updatedAt: isoDaysAgo(6),
    author: "Priya N.",
  },
  {
    id: "kb-2",
    title: "Securing Entra ID app registrations for delegated + application flows",
    category: "Security",
    tags: ["Entra ID", "OAuth", "Security"],
    problem: "A daemon (application permission) flow and an interactive (delegated permission) flow were sharing the same app registration, causing consent prompts to fail for standard users.",
    rootCause: "Application permissions require tenant-admin consent and should be isolated from delegated scopes used by interactive users, otherwise least-privilege boundaries blur and audit trails become unclear.",
    solution: "Split the workload into two app registrations: one client-credentials app scoped to application permissions for background jobs, and one public client used only for delegated, user-context calls. Apply Conditional Access and admin consent workflow to the daemon app only.",
    language: "text",
    code:
      "# Client credentials token request (daemon app)\nPOST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token\nContent-Type: application/x-www-form-urlencoded\n\nclient_id={app-id}\nscope=https://graph.microsoft.com/.default\nclient_secret={secret}\ngrant_type=client_credentials",
    updatedAt: isoDaysAgo(12),
    author: "Marcus T.",
  },
  {
    id: "kb-3",
    title: "Handling SharePoint list view threshold on large libraries",
    category: "SPFx",
    tags: ["SharePoint", "PnPjs", "Performance"],
    problem: "PnPjs queries against a library with 400k+ items throw 'The attempted operation is prohibited because it exceeds the list view threshold'.",
    rootCause: "SharePoint enforces a 5,000 item threshold on any query that isn't indexed and filtered on an indexed column, including CAML and REST/PnPjs calls under the hood.",
    solution: "Add an indexed column for the primary filter field, then page results using PnPjs' built-in paging with a filter that hits the index. Never fetch the full list unfiltered.",
    language: "typescript",
    code:
      "import { spfi } from \"@pnp/sp\";\n\nconst sp = spfi().using(SPFx(context));\n\nasync function getRecentInvoices(sp, days = 30) {\n  const cutoff = new Date();\n  cutoff.setDate(cutoff.getDate() - days);\n\n  return sp.web.lists\n    .getByTitle(\"Invoices\")\n    .items.filter(`Modified ge datetime'${cutoff.toISOString()}'`)\n    .top(500)\n    .orderBy(\"Modified\", false)();\n}",
    updatedAt: isoDaysAgo(3),
    author: "Priya N.",
  },
  {
    id: "kb-4",
    title: "Batching Microsoft Graph requests to avoid throttling",
    category: "Graph API",
    tags: ["Graph API", "Throttling", "Performance"],
    problem: "A nightly job making 200+ sequential Graph calls to sync user profile photos consistently hits 429 Too Many Requests.",
    rootCause: "Sequential per-user calls multiply round-trips and ignore Graph's batching endpoint, which is designed exactly for bulk operations and respects retry-after headers more gracefully when used correctly.",
    solution: "Group requests into batches of up to 20 using the /$batch endpoint, honor Retry-After headers on 429 responses, and add jittered exponential backoff between batches.",
    language: "json",
    code:
      "POST https://graph.microsoft.com/v1.0/$batch\nContent-Type: application/json\n\n{\n  \"requests\": [\n    { \"id\": \"1\", \"method\": \"GET\", \"url\": \"/users/alice@contoso.com/photo/$value\" },\n    { \"id\": \"2\", \"method\": \"GET\", \"url\": \"/users/bob@contoso.com/photo/$value\" }\n  ]\n}",
    updatedAt: isoDaysAgo(9),
    author: "Dana W.",
  },
  {
    id: "kb-5",
    title: "Power Automate: retry pattern for flaky HTTP connectors",
    category: "Power Platform",
    tags: ["Power Automate", "Resiliency"],
    problem: "A flow calling an on-prem gateway REST action intermittently fails with timeouts during peak hours, breaking the parent approval process.",
    rootCause: "The HTTP action's default retry policy (4 retries, exponential) isn't tuned for the gateway's actual recovery time, and failures aren't caught, so the whole flow run terminates instead of degrading gracefully.",
    solution: "Configure a custom retry policy on the action's settings (count, interval, type=exponential), then wrap the action in a Scope with a parallel 'Configure run after' failure branch that logs to the BenchIssueTracker list.",
    language: "text",
    code:
      "Retry Policy: Exponential Interval\nCount: 4\nMinimum Interval: PT5S\nMaximum Interval: PT1M\n\nRun-after configuration on downstream step:\n  has failed, has timed out, is skipped -> Log to BenchIssueTracker",
    updatedAt: isoDaysAgo(15),
    author: "Marcus T.",
  },
  {
    id: "kb-6",
    title: "SPFx bundle size creeping past recommended thresholds",
    category: "SPFx",
    tags: ["SPFx", "Webpack", "Performance"],
    problem: "gulp bundle --ship reports the main web part bundle over 3MB, causing slow first paint in the modern SharePoint page.",
    rootCause: "Moment.js with all locales and a full lodash import were pulled in instead of tree-shakeable, scoped imports, inflating the vendor chunk dramatically.",
    solution: "Replace moment with date-fns (or native Intl), switch to per-function lodash imports (lodash/get instead of lodash), and enable webpack-bundle-analyzer in the SPFx config to catch regressions in CI.",
    language: "bash",
    code:
      "npm uninstall moment lodash\nnpm install date-fns lodash.get lodash.debounce\n\n# config/config.json\n\"bundleAnalyzerConfig\": {\n  \"analyzerMode\": \"json\",\n  \"reportFilename\": \"analyzer-report.json\"\n}",
    updatedAt: isoDaysAgo(20),
    author: "Dana W.",
  },
];

const CHALLENGES = [
  {
    id: "ch-1",
    title: "Production outage: Graph subscription webhooks stopped renewing",
    severity: "Critical",
    system: "Northwind Identity Gateway",
    status: "Resolved",
    date: isoDaysAgo(8),
    rootCause: "The Azure Function responsible for renewing Graph change notification subscriptions silently failed after a certificate rotation invalidated its client secret reference in Key Vault.",
    resolutionSteps: [
      "Rotated the Key Vault secret reference and redeployed the renewal function.",
      "Manually re-subscribed all active webhooks via the Graph /subscriptions endpoint.",
      "Added a health-check ping that alerts when subscription expiry is within 6 hours.",
    ],
    resolutionTime: "3h 40m",
    preventive: "Automate certificate rotation with an Azure Monitor alert tied to Key Vault secret near-expiry events, and add renewal function to the on-call synthetic monitoring suite.",
  },
  {
    id: "ch-2",
    title: "SPFx web part fails to load in classic pages after tenant update",
    severity: "High",
    system: "Fabrikam SPFx Suite",
    status: "Resolved",
    date: isoDaysAgo(14),
    rootCause: "A SharePoint Online platform update changed the default CDN cache-control headers, serving a stale manifest that referenced a deprecated API version.",
    resolutionSteps: [
      "Bumped the solution version and forced a CDN cache purge via the tenant app catalog.",
      "Pinned the manifest's apiVersion explicitly instead of relying on 'latest'.",
      "Validated on both classic and modern pages across two test tenants.",
    ],
    resolutionTime: "5h 10m",
    preventive: "Add a scheduled canary deployment to a pre-production tenant that runs Playwright smoke tests against classic and modern page renders after every tenant update window.",
  },
  {
    id: "ch-3",
    title: "Power Automate flow silently drops approval requests",
    severity: "Medium",
    system: "Contoso Intranet Hub",
    status: "Investigating",
    date: isoDaysAgo(2),
    rootCause: "Under investigation — early signs point to the 'Start and wait for an approval' action timing out when the approver's mailbox is over quota, causing the flow to complete without notifying downstream steps.",
    resolutionSteps: [
      "Reproduced with a test mailbox intentionally set near quota.",
      "Added diagnostic logging around the approval action's status output.",
    ],
    resolutionTime: "In progress",
    preventive: "Pending root cause confirmation — likely fix is a run-after branch that escalates to a backup approver group.",
  },
  {
    id: "ch-4",
    title: "Intermittent 401s from custom API Management-fronted service",
    severity: "High",
    system: "Adventure Works Analytics",
    status: "Resolved",
    date: isoDaysAgo(21),
    rootCause: "Token caching in the SPFx client held on to expired access tokens for up to 5 minutes past actual expiry due to a clock-skew miscalculation in the custom cache wrapper.",
    resolutionSteps: [
      "Replaced the custom token cache with AadTokenProviderFactory's built-in caching.",
      "Added a 60-second safety buffer before expiry to trigger proactive refresh.",
    ],
    resolutionTime: "2h 15m",
    preventive: "Standardized token acquisition across all SPFx solutions on the built-in AadTokenProviderFactory to remove bespoke caching logic entirely.",
  },
];

const LEARNING_ITEMS = [
  { id: "l-1", type: "Certification", title: "Microsoft Certified: Power Platform Solution Architect Expert", provider: "Microsoft Learn", progress: 100, status: "Completed", dueDate: isoDaysAgo(-40), notes: "Renewed via Learn assessment; focus areas covered ALM and governance." },
  { id: "l-2", type: "Certification", title: "Microsoft Certified: Azure Developer Associate (AZ-204)", provider: "Microsoft Learn", progress: 65, status: "In Progress", dueDate: isoDaysAgo(-25), notes: "Need to review Azure Functions durable orchestration patterns before exam." },
  { id: "l-3", type: "Course", title: "Advanced TypeScript for Enterprise SPFx Solutions", provider: "Pluralsight", progress: 100, status: "Completed", dueDate: isoDaysAgo(30), notes: "Great module on generics for reusable PnPjs query builders." },
  { id: "l-4", type: "Workshop", title: "Microsoft Graph Data Connect Deep Dive", provider: "Microsoft Learn Workshop", progress: 40, status: "In Progress", dueDate: isoDaysAgo(-10), notes: "Lab 3 (Synapse pipeline integration) still pending." },
  { id: "l-5", type: "Course", title: "Power Automate: Building Resilient Enterprise Flows", provider: "Pluralsight", progress: 80, status: "In Progress", dueDate: isoDaysAgo(-5), notes: "Applying retry-policy patterns directly to the approval flow rebuild." },
  { id: "l-6", type: "Certification", title: "Microsoft Certified: Identity and Access Administrator (SC-300)", provider: "Microsoft Learn", progress: 15, status: "Planned", dueDate: isoDaysAgo(-60), notes: "Kickoff scheduled after AZ-204 exam." },
];

const COMPONENT_LIBRARY = [
  {
    id: "cl-1",
    name: "useGraphClient",
    category: "Graph API Script",
    version: "v2.3.0",
    language: "typescript",
    tags: ["React Hook", "Graph"],
    description: "React hook that memoizes an MSGraphClientV3 instance for the lifetime of a web part and exposes loading/error state.",
    code:
      "import { useEffect, useState } from \"react\";\nimport { MSGraphClientV3 } from \"@microsoft/sp-http\";\n\nexport function useGraphClient(context) {\n  const [client, setClient] = useState(null);\n  useEffect(() => {\n    let mounted = true;\n    context.msGraphClientFactory.getClient(\"3\").then((c) => {\n      if (mounted) setClient(c);\n    });\n    return () => { mounted = false; };\n  }, [context]);\n  return client;\n}",
  },
  {
    id: "cl-2",
    name: "Approval Escalation Expression",
    category: "Power Automate Expression",
    version: "v1.1.0",
    language: "text",
    tags: ["Power Automate", "Approvals"],
    description: "Expression that resolves the correct escalation approver when the primary approver's response is null after the timeout branch.",
    code:
      "if(\n  equals(outputs('Start_and_wait_for_an_approval')?['body/outcome'], null),\n  variables('BackupApproverEmail'),\n  outputs('Start_and_wait_for_an_approval')?['body/responses'][0]['responder']['email']\n)",
  },
  {
    id: "cl-3",
    name: "Patch Formula: Conditional Record Upsert",
    category: "Power Apps Formula",
    version: "v1.0.4",
    language: "text",
    tags: ["Power Apps", "Patch"],
    description: "Canvas app formula that upserts a Dataverse row only when the local form has unsaved changes, avoiding redundant writes.",
    code:
      "If(\n  EditForm1.Unsaved,\n  Patch(\n    BenchIssueTracker,\n    LookUp(BenchIssueTracker, ID = varSelectedID),\n    {\n      Status: ddStatus.Selected.Value,\n      Priority: ddPriority.Selected.Value,\n      Notes: txtNotes.Text\n    }\n  ),\n  Notify(\"No changes to save\", NotificationType.Information)\n)",
  },
  {
    id: "cl-4",
    name: "withListItemPaging",
    category: "SPFx Utility",
    version: "v3.0.1",
    language: "typescript",
    tags: ["PnPjs", "Paging"],
    description: "Utility generator that transparently pages through a PnPjs query, yielding items in chunks to keep the UI responsive on large lists.",
    code:
      "export async function* withListItemPaging(query, pageSize = 200) {\n  let page = await query.top(pageSize).getPaged();\n  yield page.results;\n  while (page.hasNext) {\n    page = await page.getNext();\n    yield page.results;\n  }\n}",
  },
  {
    id: "cl-5",
    name: "Graph Batch Builder",
    category: "Graph API Script",
    version: "v1.4.2",
    language: "typescript",
    tags: ["Graph", "Batching"],
    description: "Splits an array of Graph requests into $batch-compliant groups of 20 and executes them sequentially with backoff on 429s.",
    code:
      "export async function runGraphBatch(client, requests) {\n  const chunks = [];\n  for (let i = 0; i < requests.length; i += 20) chunks.push(requests.slice(i, i + 20));\n\n  const results = [];\n  for (const chunk of chunks) {\n    const res = await client.api(\"/$batch\").post({ requests: chunk });\n    results.push(...res.responses);\n    await new Promise((r) => setTimeout(r, 250));\n  }\n  return results;\n}",
  },
  {
    id: "cl-6",
    name: "Dynamic Approver Lookup Flow",
    category: "Power Automate Expression",
    version: "v1.0.0",
    language: "text",
    tags: ["Power Automate", "Office Scripts"],
    description: "Expression to resolve a requester's manager dynamically using the Office 365 Users connector output for approval routing.",
    code:
      "outputs('Get_manager_(V2)')?['body/mail']",
  },
];

const SEED_ISSUES = [
  { id: "is-1", title: "Kanban card drag handle unresponsive on Safari", description: "Drag events fail to register on iPadOS Safari 17 for the issue tracker board.", assignee: "Priya N.", priority: "Medium", dueDate: isoDaysAgo(-4), tags: ["Bug", "UI"], column: "todo" },
  { id: "is-2", title: "Add pagination to BenchKnowledgeBase list view", description: "List view threshold risk once KB exceeds 5,000 articles tenant-wide.", assignee: "Dana W.", priority: "Low", dueDate: isoDaysAgo(-12), tags: ["Enhancement"], column: "todo" },
  { id: "is-3", title: "Migrate legacy Power Automate flow to cloud flow v2 designer", description: "Old flow still uses deprecated 'When an item is created' trigger shape.", assignee: "Marcus T.", priority: "Medium", dueDate: isoDaysAgo(-7), tags: ["Tech Debt"], column: "inprogress" },
  { id: "is-4", title: "SPFx solution fails ALM pipeline on API permission approval step", description: "Deployment pipeline blocks waiting on manual Graph permission approval each release.", assignee: "Marcus T.", priority: "High", dueDate: isoDaysAgo(-2), tags: ["DevOps"], column: "inprogress" },
  { id: "is-5", title: "Entra ID Conditional Access blocking service principal sign-in", description: "Nightly sync job fails tenant-wide; suspected CA policy scope regression.", assignee: "Priya N.", priority: "High", dueDate: isoDaysAgo(0), tags: ["Security", "Outage-Risk"], column: "critical" },
  { id: "is-6", title: "Graph subscription webhook expiry alert missing on-call routing", description: "Alert fires but doesn't page the on-call rotation, risking a repeat of ch-1.", assignee: "Dana W.", priority: "High", dueDate: isoDaysAgo(-1), tags: ["Reliability"], column: "critical" },
];

const SP_SCHEMAS = [
  {
    name: "BenchProjects",
    description:
      "Project register — one row per project, referenced by BenchWorkLogs' ProjectName Lookup column. Column names below are display names (type them as-is when creating the column in SharePoint); Bench resolves each one's real internal name automatically at runtime, so it doesn't matter that SharePoint stores \"Project Code\" internally as something like \"Project_x0020_Code\".",
    fields: [
      { name: "Title", type: "Single line of text", description: "Project name." },
      { name: "Project Code", type: "Single line of text", description: "Short project code, e.g. CTH-001." },
      { name: "Description", type: "Multiple lines of text", description: "Free-form project summary." },
      { name: "Status", type: "Choice", description: "Planning, Active, On Hold, Completed, Cancelled." },
      { name: "Project Health", type: "Choice", description: "Green, Yellow, Red." },
      { name: "Category", type: "Choice", description: "Client, Internal, R&D." },
      { name: "Project Manager", type: "Person or Group", description: "Single owner accountable for delivery." },
      { name: "Technical Lead", type: "Person or Group", description: "Single technical owner." },
      { name: "Team Members", type: "Person or Group (multiple)", description: "Everyone staffed on the project." },
      { name: "Client / Company", type: "Single line of text", description: "Client or company name, blank for internal projects." },
      { name: "Zone", type: "Single line of text", description: "Delivery region/zone, e.g. NA-East." },
      { name: "Start Date", type: "Date and Time", description: "Project kickoff date." },
      { name: "Target End Date", type: "Date and Time", description: "Planned completion date." },
      { name: "Actual End Date", type: "Date and Time", description: "Actual completion date, blank until closed." },
      { name: "Estimated Hours", type: "Number", description: "Total estimated effort in hours." },
      { name: "Repository URL", type: "Hyperlink or Picture", description: "Link to the source repository." },
      { name: "Documentation Link", type: "Hyperlink or Picture", description: "Link to project documentation." },
    ],
  },
  {
    name: "BenchWorkLogs",
    description: "Daily work log and timesheet entries submitted by delivery team members.",
    fields: [
      { name: "Title", type: "Single line of text", description: "Task title." },
      { name: "ProjectName", type: "Lookup (BenchProjects)", description: "Associated project." },
      { name: "HoursSpent", type: "Number", description: "Hours logged, 0.25 increments." },
      { name: "Category", type: "Choice", description: "Development, Code Review, Meeting, Documentation, Testing, Deployment, Research." },
      { name: "Status", type: "Choice", description: "Completed, In Progress, Blocked." },
      { name: "GitCommitSignature", type: "Single line of text", description: "Optional short SHA reference." },
      { name: "Jira Ticket Link", type: "Hyperlink or Picture", description: "Optional link to the related Jira ticket." },
      { name: "EntryDate", type: "Date and Time", description: "Indexed for list view threshold performance." },
      { name: "SubmittedBy", type: "Person or Group", description: "Auto-populated from context.pageContext.user." },
    ],
  },
  {
    name: "BenchKnowledgeBase",
    description: "Searchable technical knowledge repository articles.",
    fields: [
      { name: "Title", type: "Single line of text", description: "Article headline." },
      { name: "CategoryTag", type: "Choice", description: "Integrations, Security, SPFx, Graph API, Power Platform." },
      { name: "Problem", type: "Multiple lines of text", description: "Rich text problem statement." },
      { name: "RootCause", type: "Multiple lines of text", description: "Rich text root cause narrative." },
      { name: "Solution", type: "Multiple lines of text", description: "Rich text resolution guidance." },
      { name: "CodeSnippet", type: "Multiple lines of text (plain)", description: "Raw code block, rendered with syntax highlighting." },
      { name: "Language", type: "Choice", description: "typescript, json, bash, text, etc." },
      { name: "Author", type: "Person or Group", description: "Original author of the article." },
    ],
  },
  {
    name: "BenchChallenges",
    description: "Technical incident and challenge resolution log.",
    fields: [
      { name: "Title", type: "Single line of text", description: "Incident summary." },
      { name: "Severity", type: "Choice", description: "Critical, High, Medium." },
      { name: "AffectedSystem", type: "Lookup (BenchProjects)", description: "Impacted project or system." },
      { name: "RootCauseAnalysis", type: "Multiple lines of text", description: "Detailed RCA narrative." },
      { name: "ResolutionSteps", type: "Multiple lines of text", description: "Newline-delimited remediation steps." },
      { name: "ResolutionTime", type: "Single line of text", description: "Elapsed time to resolution." },
      { name: "PreventiveRecommendation", type: "Multiple lines of text", description: "Follow-up hardening actions." },
    ],
  },
  {
    name: "BenchLearningCenter",
    description: "Certification, course, and workshop progress tracking.",
    fields: [
      { name: "Title", type: "Single line of text", description: "Certification / course / workshop name." },
      { name: "ItemType", type: "Choice", description: "Certification, Course, Workshop." },
      { name: "Provider", type: "Single line of text", description: "Microsoft Learn, Pluralsight, etc." },
      { name: "ProgressPercent", type: "Number", description: "0-100 completion." },
      { name: "Status", type: "Choice", description: "Planned, In Progress, Completed." },
      { name: "StudyNotes", type: "Multiple lines of text", description: "Free-form notes." },
      { name: "TargetDate", type: "Date and Time", description: "Exam or completion target." },
    ],
  },
  {
    name: "BenchProjectDocs",
    description: "Project-level documentation and architecture references.",
    fields: [
      { name: "Title", type: "Single line of text", description: "Document title." },
      { name: "ProjectName", type: "Lookup (BenchProjects)", description: "Owning project." },
      { name: "DocumentType", type: "Choice", description: "Architecture, Runbook, Onboarding, Retrospective." },
      { name: "FileRef", type: "Hyperlink or Picture", description: "Link to the document library item." },
      { name: "Owner", type: "Person or Group", description: "Document owner." },
    ],
  },
  {
    name: "BenchReusableComponents",
    description: "Boilerplate library of reusable SPFx, Power Platform, and Graph code assets.",
    fields: [
      { name: "Title", type: "Single line of text", description: "Component name." },
      { name: "ComponentCategory", type: "Choice", description: "SPFx Utility, Power Apps Formula, Power Automate Expression, Graph API Script." },
      { name: "Version", type: "Single line of text", description: "Semver string." },
      { name: "Description", type: "Multiple lines of text", description: "Usage summary." },
      { name: "CodeBody", type: "Multiple lines of text (plain)", description: "Source snippet." },
      { name: "Tags", type: "Managed Metadata", description: "Taxonomy tags for discovery." },
    ],
  },
  {
    name: "BenchIssueTracker",
    description: "Kanban-backed issues, risks, and bug tracking list.",
    fields: [
      { name: "Title", type: "Single line of text", description: "Ticket summary." },
      { name: "Description", type: "Multiple lines of text", description: "Detailed ticket body." },
      { name: "AssignedTo", type: "Person or Group", description: "Ticket owner." },
      { name: "Priority", type: "Choice", description: "Low, Medium, High." },
      { name: "DueDate", type: "Date and Time", description: "Target resolution date." },
      { name: "BoardColumn", type: "Choice", description: "To Do, In Progress, Critical / At Risk." },
      { name: "Tags", type: "Managed Metadata", description: "Bug, Enhancement, Tech Debt, Security, etc." },
    ],
  },
];

const SPFX_TREE = `bench-spfx-solution/
├── config/
│   ├── config.json
│   ├── package-solution.json
│   └── serve.json
├── sharepoint/
│   └── assets/
├── src/
│   ├── webparts/
│   │   └── benchDashboard/
│   │       ├── components/
│   │       │   ├── BenchDashboard.tsx
│   │       │   ├── BenchDashboard.module.scss
│   │       │   └── IBenchDashboardProps.ts
│   │       ├── loc/
│   │       ├── BenchDashboardWebPart.manifest.json
│   │       └── BenchDashboardWebPart.ts
│   ├── services/
│   │   ├── GraphService.ts
│   │   ├── SPDataService.ts
│   │   └── TokenProvider.ts
│   ├── models/
│   │   └── IWorkLogItem.ts
│   ├── hooks/
│   │   └── useGraphClient.ts
│   └── extensions/
│       └── benchHeaderCommandSet/
├── teams/
├── gulpfile.js
├── package.json
└── tsconfig.json`;

const POWER_AUTOMATE_FLOWS = [
  {
    id: "pa-1",
    name: "Daily Work Log Digest",
    trigger: "Recurrence — every weekday at 6:00 PM",
    actions: ["Get items (BenchWorkLogs, filtered by today)", "Compose HTML table of entries per team member", "Send an email (V2) digest to delivery leads", "Post adaptive card summary to Teams channel"],
    description: "Aggregates the day's submitted work logs and pushes a summary digest to project leads and the delivery Teams channel.",
  },
  {
    id: "pa-2",
    name: "Critical Challenge Escalation",
    trigger: "When an item is created or modified — BenchChallenges (Severity = Critical)",
    actions: ["Start and wait for an approval (incident commander acknowledgement)", "Create a Planner task in the Incident Response plan", "Post to #incident-bridge Teams channel", "Send SMS/voice call via Power Automate approvals mobile notification"],
    description: "Fires an immediate multi-channel escalation whenever a Critical-severity challenge is logged.",
  },
  {
    id: "pa-3",
    name: "Certification Renewal Reminder",
    trigger: "Recurrence — weekly, Monday 8:00 AM",
    actions: ["Get items (BenchLearningCenter, TargetDate within 30 days)", "Filter array by Status ne Completed", "Send reminder email with study resources", "Update reminder-sent flag on the item"],
    description: "Scans upcoming certification target dates and reminds owners with 30-day lead time.",
  },
  {
    id: "pa-4",
    name: "Issue Tracker Auto-Triage",
    trigger: "When an item is created — BenchIssueTracker",
    actions: ["Condition: Tags contains 'Security' or 'Outage-Risk'", "If yes: set BoardColumn = Critical / At Risk, notify security channel", "If no: run AI Builder priority classification, set Priority accordingly"],
    description: "Automatically routes newly created tickets into the correct kanban column based on tags and AI-assisted priority classification.",
  },
];

const ANNOUNCEMENTS = [
  { id: "an-1", type: "security", title: "Mandatory MFA re-enrollment window closes Friday", body: "All delivery team members must re-enroll authenticator methods under the updated Conditional Access policy before Friday 5 PM.", date: isoDaysAgo(1) },
  { id: "an-2", type: "info", title: "New SPFx 1.19 baseline for all active solutions", body: "Please upgrade local dev environments and validate gulp bundle --ship against SharePoint Online before next sprint's deployment window.", date: isoDaysAgo(2) },
  { id: "an-3", type: "success", title: "Northwind Identity Gateway passed security review", body: "No critical findings. Two medium recommendations have been logged to the Issue Tracker for follow-up.", date: isoDaysAgo(4) },
];

const QUICK_LAUNCH = [
  { id: "ql-1", label: "Azure Portal", href: "https://portal.azure.com", icon: Cloud },
  { id: "ql-2", label: "M365 Admin Center", href: "https://admin.microsoft.com", icon: Users },
  { id: "ql-3", label: "SPFx SDK Docs", href: "https://learn.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview", icon: FileCode2 },
  { id: "ql-4", label: "SharePoint List Schemas", view: "admin", icon: Database },
];

const CLIPPY_QUICK_PROMPTS = [
  { label: "Summarize Work Logs", icon: ClipboardList },
  { label: "Power Automate Expression", icon: Workflow },
  { label: "Explain Graph Batching", icon: Database },
  { label: "Draft Resolution Notes", icon: AlertTriangle },
];

const AUTH_PROFILES = {
  google: { name: "Alex Rivera", email: "alex.rivera@gmail.com", role: "Senior Consultant", provider: "google" },
  microsoft: { name: "Alex Rivera", email: "alex.rivera@contoso.com", role: "Senior Consultant", provider: "microsoft" },
};

const NAV_ITEMS = [
  { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { key: "worklog", label: "Work Log & Timesheets", icon: ClipboardList },
  { key: "projects", label: "Projects", icon: Briefcase },
  { key: "knowledge", label: "Knowledge Base", icon: BookOpen },
  { key: "challenges", label: "Technical Challenges", icon: AlertTriangle },
  { key: "learning", label: "Learning Center", icon: GraduationCap },
  { key: "components", label: "Components Library", icon: Code2 },
  { key: "issues", label: "Issues & Risks", icon: KanbanSquare },
  { key: "admin", label: "Architecture (Admin)", icon: Settings2 },
];

/* ----------------------------------------------------------------------- */
/* Clippy AI reply engine                                                   */
/* ----------------------------------------------------------------------- */

function generateClippyReply(rawText, ctx) {
  const text = rawText.toLowerCase();

  if (text.includes("summarize") && text.includes("work log")) {
    const totalHours = ctx.workLogs.reduce((sum, l) => sum + Number(l.hours || 0), 0);
    const byCategory = {};
    ctx.workLogs.forEach((l) => {
      byCategory[l.category] = (byCategory[l.category] || 0) + Number(l.hours || 0);
    });
    const breakdown = Object.entries(byCategory)
      .sort((a, b) => b[1] - a[1])
      .map(([cat, hrs]) => `• ${cat}: ${hrs.toFixed(1)}h`)
      .join("\n");
    return `Here's your work log summary across ${ctx.workLogs.length} entries:\n\nTotal hours logged: ${totalHours.toFixed(1)}h\n\nBreakdown by category:\n${breakdown}\n\nWant me to draft a status update for your project lead?`;
  }

  if (text.includes("power automate") && text.includes("expression")) {
    return `Here's a commonly useful Power Automate expression pattern for dynamic approver routing:\n\noutputs('Get_manager_(V2)')?['body/mail']\n\nAnd a safe null-coalescing pattern for optional fields:\n\ncoalesce(triggerBody()?['field_1'], 'Not provided')\n\nTip: always use the ?[] safe-navigation syntax when reading connector outputs — it prevents the whole flow run from failing on a missing property.`;
  }

  if (text.includes("graph") && text.includes("batch")) {
    return `Microsoft Graph batching lets you combine up to 20 requests into a single HTTP call to /$batch, cutting round-trips dramatically.\n\nPOST https://graph.microsoft.com/v1.0/$batch\n{\n  "requests": [\n    { "id": "1", "method": "GET", "url": "/me" },\n    { "id": "2", "method": "GET", "url": "/me/messages?$top=5" }\n  ]\n}\n\nEach sub-request runs independently and returns its own status code, so a single failure in the batch won't fail the others. Respect Retry-After headers on any 429 responses within the batch.`;
  }

  if (text.includes("resolution") || text.includes("rca") || text.includes("root cause")) {
    const critical = ctx.challenges.filter((c) => c.severity === "Critical").length;
    return `I can help draft resolution notes. Structure I'd recommend:\n\n1. Impact summary (system, duration, users affected)\n2. Root cause (technical, one sentence)\n3. Immediate remediation steps taken\n4. Preventive follow-up items (link to Issue Tracker)\n\nYou currently have ${critical} Critical-severity challenge${critical === 1 ? "" : "s"} logged — want me to pull the most recent one as a starting template?`;
  }

  if (text.includes("certification") || text.includes("learning") || text.includes("exam")) {
    return `Your Learning Center shows a healthy mix in progress. Keep momentum by time-boxing 3-4 focused hours weekly against your nearest target date item — those tend to slip first when sprint work spikes.`;
  }

  if (text.includes("hello") || text.includes("hi") || text.trim() === "") {
    return `Hi, I'm Clippy, your Bench AI Copilot! I can summarize your work logs, explain SPFx/Graph/Power Platform patterns, or help draft resolution notes. What would you like help with?`;
  }

  return `Good question! Based on your Bench workspace data, here's a quick take: keep integrations behind the AadHttpClient/MSGraphClient factories for auth safety, batch bulk Graph calls, and log anything Critical straight into the Issue Tracker so the escalation flow picks it up automatically. Want more detail on any of those?`;
}

/* ----------------------------------------------------------------------- */
/* Small UI primitives                                                      */
/* ----------------------------------------------------------------------- */

function Badge({ children, tone = "slate" }) {
  const tones = {
    slate: "bg-slate-100 text-slate-700 dark:bg-slate-800 dark:text-slate-300 ring-slate-300/60 dark:ring-slate-700",
    blue: "bg-blue-50 text-blue-700 dark:bg-blue-500/10 dark:text-blue-300 ring-blue-300/60 dark:ring-blue-500/30",
    green: "bg-emerald-50 text-emerald-700 dark:bg-emerald-500/10 dark:text-emerald-300 ring-emerald-300/60 dark:ring-emerald-500/30",
    red: "bg-red-50 text-red-700 dark:bg-red-500/10 dark:text-red-300 ring-red-300/60 dark:ring-red-500/30",
    orange: "bg-orange-50 text-orange-700 dark:bg-orange-500/10 dark:text-orange-300 ring-orange-300/60 dark:ring-orange-500/30",
    yellow: "bg-yellow-50 text-yellow-800 dark:bg-yellow-500/10 dark:text-yellow-300 ring-yellow-300/60 dark:ring-yellow-500/30",
    purple: "bg-purple-50 text-purple-700 dark:bg-purple-500/10 dark:text-purple-300 ring-purple-300/60 dark:ring-purple-500/30",
  };
  return (
    <span className={cx("inline-flex items-center gap-1 rounded-full px-2.5 py-0.5 text-xs font-medium ring-1 ring-inset", tones[tone] || tones.slate)}>
      {children}
    </span>
  );
}

function Card({ children, className = "" }) {
  return (
    <div
      className={cx(
        "rounded-2xl border border-slate-200/80 bg-white/80 shadow-sm backdrop-blur-xl transition-colors duration-300",
        "dark:border-slate-800/80 dark:bg-slate-900/60",
        className
      )}
    >
      {children}
    </div>
  );
}

function IconButton({ onClick, title, children, className = "" }) {
  return (
    <button
      type="button"
      onClick={onClick}
      title={title}
      className={cx(
        "inline-flex h-9 w-9 items-center justify-center rounded-xl border border-slate-200 bg-white text-slate-600 transition hover:bg-slate-50 hover:text-slate-900",
        "dark:border-slate-700 dark:bg-slate-800 dark:text-slate-300 dark:hover:bg-slate-700 dark:hover:text-white",
        className
      )}
    >
      {children}
    </button>
  );
}

function ProgressBar({ value, tone = "blue" }) {
  const tones = {
    blue: "bg-blue-600",
    green: "bg-emerald-500",
    orange: "bg-orange-500",
    purple: "bg-purple-500",
  };
  return (
    <div className="h-2 w-full overflow-hidden rounded-full bg-slate-100 dark:bg-slate-800">
      <div
        className={cx("h-full rounded-full transition-all duration-500", tones[tone] || tones.blue)}
        style={{ width: `${Math.min(100, Math.max(0, value))}%` }}
      />
    </div>
  );
}

function Modal({ open, onClose, title, children, wide = false, footer = null }) {
  useEffect(() => {
    if (!open) return;
    const handleKey = (e) => e.key === "Escape" && onClose();
    window.addEventListener("keydown", handleKey);
    return () => window.removeEventListener("keydown", handleKey);
  }, [open, onClose]);

  if (!open) return null;
  return (
    // Modal is always rendered as a sibling inside some view's `space-y-6`
    // wrapper, whose "> * ~ *" selector adds a margin-top to every sibling
    // after the first — including this one, despite it being fixed/off-flow.
    // `!m-0` force-overrides that regardless of which one has higher
    // selector specificity, so the overlay always starts flush at the top
    // of the viewport instead of leaving a gap that exposes the header.
    <div className="fixed inset-0 z-[60] !m-0 flex items-center justify-center p-4">
      <div className="absolute inset-0 bg-slate-900/50 backdrop-blur-sm animate-fade-in" onClick={onClose} />
      <div
        className={cx(
          "relative z-10 flex max-h-[85vh] w-full flex-col overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-2xl animate-pop-in",
          "dark:border-slate-800 dark:bg-slate-900",
          wide ? "max-w-3xl" : "max-w-lg"
        )}
      >
        <div className="flex shrink-0 items-start justify-between gap-4 border-b border-slate-100 px-6 py-4 dark:border-slate-800">
          <h3 className="text-lg font-semibold text-slate-900 dark:text-white">{title}</h3>
          <button
            onClick={onClose}
            title="Close"
            className="rounded-lg p-1 text-slate-400 transition hover:bg-slate-100 hover:text-slate-700 dark:hover:bg-slate-800 dark:hover:text-slate-200"
          >
            <X size={18} />
          </button>
        </div>
        <div className="min-h-0 flex-1 overflow-y-auto px-6 py-5">{children}</div>
        {footer && <div className="shrink-0 border-t border-slate-100 px-6 py-4 dark:border-slate-800">{footer}</div>}
      </div>
    </div>
  );
}

function CodeBlock({ code, language, copyKey, copiedKey, onCopy }) {
  return (
    <div className="relative overflow-hidden rounded-xl border border-slate-800 bg-slate-950">
      <div className="flex items-center justify-between border-b border-slate-800 px-4 py-2">
        <span className="text-xs font-medium uppercase tracking-wide text-slate-400">{language}</span>
        <button
          onClick={() => onCopy(code, copyKey)}
          className="inline-flex items-center gap-1.5 rounded-lg px-2 py-1 text-xs font-medium text-slate-300 transition hover:bg-slate-800 hover:text-white"
        >
          {copiedKey === copyKey ? (
            <>
              <Check size={13} className="text-emerald-400" /> Copied
            </>
          ) : (
            <>
              <Copy size={13} /> Copy
            </>
          )}
        </button>
      </div>
      <pre className="overflow-x-auto p-4 text-sm leading-relaxed text-slate-200">
        <code>{code}</code>
      </pre>
    </div>
  );
}

function EmptyState({ icon: Icon, message }) {
  return (
    <div className="flex flex-col items-center justify-center gap-2 rounded-2xl border border-dashed border-slate-300 py-14 text-center text-slate-400 dark:border-slate-700">
      <Icon size={28} />
      <p className="text-sm">{message}</p>
    </div>
  );
}

function SyncBadge({ status, error }) {
  const config = {
    unconfigured: { icon: CloudOff, tone: "slate", label: "Local demo data" },
    idle: { icon: CloudOff, tone: "slate", label: "Local demo data" },
    connecting: { icon: RefreshCw, tone: "blue", label: "Syncing with SharePoint…", spin: true },
    connected: { icon: CheckCircle2, tone: "green", label: "Synced with SharePoint" },
    error: { icon: AlertCircle, tone: "red", label: "Sync error — showing local data" },
  }[status || "unconfigured"];

  return (
    <span title={status === "error" ? error : undefined}>
      <Badge tone={config.tone}>
        <config.icon size={11} className={config.spin ? "animate-spin" : undefined} />
        {config.label}
      </Badge>
    </span>
  );
}

/* ----------------------------------------------------------------------- */
/* Clippy assistant                                                         */
/* ----------------------------------------------------------------------- */

function ClippyAvatar({ mood = "idle" }) {
  const animClass =
    mood === "thinking" ? "animate-clippy-think" : mood === "celebrate" ? "animate-clippy-bounce" : "animate-clippy-float";
  return (
    <svg
      viewBox="0 0 120 140"
      className={cx("h-16 w-16 drop-shadow-xl", animClass)}
      style={{ transformOrigin: "50% 90%" }}
    >
      <defs>
        <linearGradient id="clipMetal" x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stopColor="#e2e8f0" />
          <stop offset="45%" stopColor="#94a3b8" />
          <stop offset="100%" stopColor="#475569" />
        </linearGradient>
        <linearGradient id="clipShine" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stopColor="#ffffff" stopOpacity="0.9" />
          <stop offset="100%" stopColor="#ffffff" stopOpacity="0" />
        </linearGradient>
      </defs>

      {/* paperclip body */}
      <path
        d="M40 20 C40 10, 55 5, 65 12 L65 95 C65 108, 50 112, 42 103 C36 96, 38 88, 46 85 L46 30 C46 24, 54 22, 56 28 L56 88"
        fill="none"
        stroke="url(#clipMetal)"
        strokeWidth="9"
        strokeLinecap="round"
      />
      <path
        d="M40 20 C40 10, 55 5, 65 12 L65 95 C65 108, 50 112, 42 103"
        fill="none"
        stroke="url(#clipShine)"
        strokeWidth="2.5"
        strokeLinecap="round"
        opacity="0.7"
      />

      {/* eyebrows */}
      <g className="animate-clippy-eyebrow" style={{ transformOrigin: "48px 55px" }}>
        <rect x="42" y="53" width="12" height="3" rx="1.5" fill="#1e293b" />
      </g>
      <g className="animate-clippy-eyebrow" style={{ transformOrigin: "68px 55px", animationDelay: "0.15s" }}>
        <rect x="62" y="53" width="12" height="3" rx="1.5" fill="#1e293b" />
      </g>

      {/* eyes */}
      <g className="animate-clippy-blink" style={{ transformOrigin: "48px 64px" }}>
        <ellipse cx="48" cy="64" rx="5.5" ry="7" fill="#ffffff" stroke="#1e293b" strokeWidth="1.5" />
        <circle cx="49" cy="65" r="2.4" fill="#1e293b" />
      </g>
      <g className="animate-clippy-blink" style={{ transformOrigin: "68px 64px", animationDelay: "0.1s" }}>
        <ellipse cx="68" cy="64" rx="5.5" ry="7" fill="#ffffff" stroke="#1e293b" strokeWidth="1.5" />
        <circle cx="69" cy="65" r="2.4" fill="#1e293b" />
      </g>
    </svg>
  );
}

function ClippyWidget({ open, onToggle, mood, showHint }) {
  return (
    <div className="fixed bottom-6 right-6 z-50 flex flex-col items-end gap-2">
      {showHint && !open && (
        <div className="relative mr-2 max-w-[190px] animate-pop-in rounded-2xl border-2 border-slate-800 bg-[#FFFFE1] px-3 py-2 text-xs font-medium text-slate-800 shadow-lg">
          Need a hand? Click me for AI help across Bench!
          <span className="absolute -bottom-2 right-6 h-3 w-3 rotate-45 border-b-2 border-r-2 border-slate-800 bg-[#FFFFE1]" />
        </div>
      )}
      <button
        type="button"
        onClick={onToggle}
        title="Bench AI Copilot"
        className="group relative flex h-20 w-20 items-center justify-center rounded-full bg-gradient-to-br from-white/90 to-slate-100/90 shadow-2xl ring-4 ring-white/60 transition hover:scale-105 dark:from-slate-800/90 dark:to-slate-900/90 dark:ring-slate-700/60"
      >
        <ClippyAvatar mood={mood} />
        <span className="absolute -right-1 -top-1 flex h-5 w-5 items-center justify-center rounded-full bg-emerald-500 text-white ring-2 ring-white dark:ring-slate-900">
          <Sparkles size={11} />
        </span>
      </button>
    </div>
  );
}

function ChatDrawer({ open, onClose, messages, onSend, mood }) {
  const [input, setInput] = useState("");
  const scrollRef = useRef(null);

  useEffect(() => {
    if (scrollRef.current) scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
  }, [messages, open]);

  if (!open) return null;

  const submit = (text) => {
    const value = (text ?? input).trim();
    if (!value) return;
    onSend(value);
    setInput("");
  };

  return (
    <div className="fixed bottom-28 right-6 z-50 flex w-[calc(100vw-3rem)] max-w-sm animate-pop-in flex-col overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-2xl dark:border-slate-800 dark:bg-slate-900 sm:w-96">
      <div className="flex items-center justify-between gap-2 border-b border-slate-200 bg-gradient-to-r from-blue-600 to-indigo-600 px-4 py-3 dark:border-slate-800">
        <div className="flex items-center gap-2 text-white">
          <Sparkles size={16} />
          <div>
            <p className="text-sm font-semibold leading-none">Bench AI Copilot</p>
            <p className="mt-0.5 text-[11px] text-blue-100">{mood === "thinking" ? "Thinking…" : "Grounded in your workspace"}</p>
          </div>
        </div>
        <button onClick={onClose} className="rounded-lg p-1 text-blue-100 transition hover:bg-white/10 hover:text-white">
          <X size={16} />
        </button>
      </div>

      <div ref={scrollRef} className="flex max-h-96 min-h-[16rem] flex-col gap-3 overflow-y-auto px-4 py-4">
        {messages.map((m) => (
          <div key={m.id} className={cx("flex", m.role === "user" ? "justify-end" : "justify-start")}>
            <div
              className={cx(
                "max-w-[85%] whitespace-pre-line rounded-2xl px-3.5 py-2.5 text-sm leading-relaxed shadow-sm",
                m.role === "user"
                  ? "rounded-br-sm bg-blue-600 text-white"
                  : "rounded-bl-sm border border-slate-200 bg-slate-50 text-slate-700 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200"
              )}
            >
              {m.text}
            </div>
          </div>
        ))}
        {mood === "thinking" && (
          <div className="flex justify-start">
            <div className="flex items-center gap-2 rounded-2xl rounded-bl-sm border border-slate-200 bg-slate-50 px-3.5 py-2.5 text-sm text-slate-500 dark:border-slate-700 dark:bg-slate-800">
              <Loader2 size={14} className="animate-spin" /> Clippy is thinking…
            </div>
          </div>
        )}
      </div>

      <div className="border-t border-slate-200 px-3 py-2 dark:border-slate-800">
        <div className="mb-2 flex flex-wrap gap-1.5">
          {CLIPPY_QUICK_PROMPTS.map((qp) => (
            <button
              key={qp.label}
              onClick={() => submit(qp.label)}
              className="inline-flex items-center gap-1 rounded-full border border-slate-200 bg-slate-50 px-2.5 py-1 text-[11px] font-medium text-slate-600 transition hover:border-blue-300 hover:bg-blue-50 hover:text-blue-700 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-300 dark:hover:border-blue-500/40 dark:hover:bg-blue-500/10 dark:hover:text-blue-300"
            >
              <qp.icon size={11} /> {qp.label}
            </button>
          ))}
        </div>
        <div className="flex items-center gap-2">
          <input
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && submit()}
            placeholder="Ask Clippy about your work…"
            className="flex-1 rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm text-slate-800 outline-none transition focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:text-slate-100 dark:focus:bg-slate-900"
          />
          <button
            onClick={() => submit()}
            className="inline-flex h-9 w-9 items-center justify-center rounded-xl bg-blue-600 text-white transition hover:bg-blue-700"
          >
            <Send size={15} />
          </button>
        </div>
      </div>
    </div>
  );
}

/* ----------------------------------------------------------------------- */
/* Authentication                                                           */
/* ----------------------------------------------------------------------- */

function GoogleIcon({ size = 18 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 48 48" aria-hidden="true">
      <path fill="#FFC107" d="M43.611,20.083H42V20H24v8h11.303c-1.649,4.657-6.08,8-11.303,8c-6.627,0-12-5.373-12-12c0-6.627,5.373-12,12-12c3.059,0,5.842,1.154,7.961,3.039l5.657-5.657C34.046,6.053,29.268,4,24,4C12.955,4,4,12.955,4,24c0,11.045,8.955,20,20,20c11.045,0,20-8.955,20-20C44,22.659,43.862,21.35,43.611,20.083z" />
      <path fill="#FF3D00" d="M6.306,14.691l6.571,4.819C14.655,15.108,18.961,12,24,12c3.059,0,5.842,1.154,7.961,3.039l5.657-5.657C34.046,6.053,29.268,4,24,4C16.318,4,9.656,8.337,6.306,14.691z" />
      <path fill="#4CAF50" d="M24,44c5.166,0,9.86-1.977,13.409-5.192l-6.19-5.238C29.211,35.091,26.715,36,24,36c-5.202,0-9.619-3.317-11.283-7.946l-6.522,5.025C9.505,39.556,16.227,44,24,44z" />
      <path fill="#1976D2" d="M43.611,20.083H42V20H24v8h11.303c-0.792,2.237-2.231,4.166-4.087,5.571c0.001-0.001,0.002-0.001,0.003-0.002l6.19,5.238C36.971,39.205,44,34,44,24C44,22.659,43.862,21.35,43.611,20.083z" />
    </svg>
  );
}

function MicrosoftIcon({ size = 18 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 23 23" aria-hidden="true">
      <rect x="1" y="1" width="10" height="10" fill="#f25022" />
      <rect x="12" y="1" width="10" height="10" fill="#7fba00" />
      <rect x="1" y="12" width="10" height="10" fill="#00a4ef" />
      <rect x="12" y="12" width="10" height="10" fill="#ffb900" />
    </svg>
  );
}

function titleCaseFromEmail(email) {
  const local = email.split("@")[0] || "";
  const words = local.replace(/[._-]+/g, " ").trim();
  if (!words) return "Bench User";
  return words
    .split(" ")
    .map((w) => (w ? w[0].toUpperCase() + w.slice(1) : w))
    .join(" ");
}

function LoginView({
  mode,
  onSsoLogin,
  ssoLoading,
  ssoError,
  onEmailLogin,
  onEmailSignup,
  formLoading,
  authError,
  onGuestLogin,
  theme,
  onToggleTheme,
}) {
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [error, setError] = useState("");

  const anyLoading = formLoading || !!ssoLoading;

  const submit = (e) => {
    e.preventDefault();
    if (!email.trim()) {
      setError("Enter your email to continue.");
      return;
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim())) {
      setError("Enter a valid email address.");
      return;
    }
    if (mode === "signup" && !name.trim()) {
      setError("Enter your full name to create an account.");
      return;
    }
    setError("");
    const payload = {
      name: mode === "signup" ? name.trim() : titleCaseFromEmail(email.trim()),
      email: email.trim(),
    };
    if (mode === "signup") {
      onEmailSignup(payload);
    } else {
      onEmailLogin(payload);
    }
  };

  return (
    <div className="relative flex min-h-screen items-center justify-center overflow-hidden bg-slate-100 px-4 py-10 transition-colors duration-300 dark:bg-slate-950">
      <div className="pointer-events-none absolute inset-0 overflow-hidden">
        <div className="absolute -left-24 -top-24 h-72 w-72 rounded-full bg-blue-400/20 blur-3xl" />
        <div className="absolute -right-24 -bottom-24 h-80 w-80 rounded-full bg-purple-400/20 blur-3xl" />
      </div>

      <button
        onClick={onToggleTheme}
        title="Toggle theme"
        className="absolute right-5 top-5 inline-flex h-9 w-9 items-center justify-center rounded-xl border border-slate-200 bg-white text-slate-600 transition hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-300"
      >
        {theme === "dark" ? <Sun size={16} /> : <Moon size={16} />}
      </button>

      <Card className="relative z-10 w-full max-w-sm p-8">
        <div className="mb-6 flex flex-col items-center text-center">
          <div className="mb-3 flex h-12 w-12 items-center justify-center rounded-2xl bg-gradient-to-br from-blue-600 to-indigo-600 text-white shadow-md">
            <Layers size={22} />
          </div>
          <h1 className="text-xl font-bold text-slate-900 dark:text-white">Bench</h1>
          <p className="mt-1 text-sm text-slate-500 dark:text-slate-400">Team Productivity &amp; Knowledge Portal</p>
        </div>

        {/* Sign Up is disabled for now — see manualAuth.js. Sign In alone
            checks the entered email against BenchAdministration. */}
        <div className="mb-5 rounded-xl bg-slate-100 p-1 text-center text-sm font-semibold text-slate-900 dark:bg-slate-800 dark:text-white">
          <span className="block rounded-lg bg-white py-1.5 shadow-sm dark:bg-slate-700">Sign In</span>
        </div>

        <form onSubmit={submit} className="space-y-3">
          {mode === "signup" && (
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Full Name</label>
              <input
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="Alex Rivera"
                className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2.5 text-sm outline-none transition focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
              />
            </div>
          )}

          <div>
            <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Email</label>
            <div className="relative">
              <Mail size={15} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="alex.rivera@contoso.com"
                autoComplete="email"
                className="w-full rounded-xl border border-slate-200 bg-slate-50 py-2.5 pl-9 pr-3 text-sm outline-none transition focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
              />
            </div>
          </div>

          {(error || authError) && (
            <p className="rounded-lg bg-red-50 px-3 py-2 text-xs font-medium text-red-600 dark:bg-red-500/10 dark:text-red-400">
              {error || authError}
            </p>
          )}

          <button
            type="submit"
            disabled={anyLoading}
            className="flex w-full items-center justify-center gap-2 rounded-xl bg-blue-600 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-700 disabled:cursor-not-allowed disabled:opacity-60"
          >
            {formLoading ? (
              <>
                <Loader2 size={15} className="animate-spin" /> {mode === "signin" ? "Signing in…" : "Creating account…"}
              </>
            ) : (
              <>
                {mode === "signin" ? "Sign In" : "Create Account"} <ArrowRight size={15} />
              </>
            )}
          </button>
        </form>

        <div className="my-5 flex items-center gap-3">
          <div className="h-px flex-1 bg-slate-200 dark:bg-slate-800" />
          <span className="text-xs font-medium text-slate-400">or continue with</span>
          <div className="h-px flex-1 bg-slate-200 dark:bg-slate-800" />
        </div>

        <div className="space-y-3">
          <button
            type="button"
            onClick={() => onSsoLogin("microsoft")}
            disabled={anyLoading}
            className="flex w-full items-center justify-center gap-3 rounded-xl border border-slate-200 bg-white px-4 py-2.5 text-sm font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200 dark:hover:bg-slate-700"
          >
            {ssoLoading === "microsoft" ? <Loader2 size={16} className="animate-spin text-slate-400" /> : <MicrosoftIcon />}
            {ssoLoading === "microsoft" ? "Connecting to Microsoft…" : "Microsoft"}
          </button>

          <button
            type="button"
            onClick={() => onSsoLogin("google")}
            disabled={anyLoading}
            className="flex w-full items-center justify-center gap-3 rounded-xl border border-slate-200 bg-white px-4 py-2.5 text-sm font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200 dark:hover:bg-slate-700"
          >
            {ssoLoading === "google" ? <Loader2 size={16} className="animate-spin text-slate-400" /> : <GoogleIcon />}
            {ssoLoading === "google" ? "Connecting to Google…" : "Google"}
          </button>
        </div>

        {ssoError && (
          <p className="mt-3 rounded-lg bg-red-50 px-3 py-2 text-xs font-medium text-red-600 dark:bg-red-500/10 dark:text-red-400">{ssoError}</p>
        )}

        <button
          type="button"
          onClick={onGuestLogin}
          className="mt-4 w-full rounded-xl border border-dashed border-slate-300 px-4 py-2 text-xs font-semibold text-slate-500 transition hover:border-slate-400 hover:text-slate-700 dark:border-slate-700 dark:text-slate-400 dark:hover:border-slate-600 dark:hover:text-slate-200"
        >
          Skip sign-in — explore as a guest
        </button>

        <p className="mt-6 text-center text-xs text-slate-400">
          Protected by Microsoft Entra ID conditional access. By continuing you agree to Bench's internal usage policy.
        </p>
      </Card>
    </div>
  );
}

/* ----------------------------------------------------------------------- */
/* Layout: Header + Sidebar                                                 */
/* ----------------------------------------------------------------------- */

function Header({ theme, onToggleTheme, onToggleMobileMenu, search, onSearch, user, onLogout, accessLevel }) {
  const [menuOpen, setMenuOpen] = useState(false);
  const initials = user.name
    .split(" ")
    .map((n) => n[0])
    .join("")
    .slice(0, 2)
    .toUpperCase();

  return (
    <header className="sticky top-0 z-40 border-b border-slate-200/80 bg-white/70 backdrop-blur-xl transition-colors duration-300 dark:border-slate-800/80 dark:bg-slate-950/70">
      <div className="flex h-16 items-center gap-3 px-4 sm:px-6">
        <button
          onClick={onToggleMobileMenu}
          className="inline-flex h-9 w-9 items-center justify-center rounded-xl text-slate-600 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800 md:hidden"
        >
          <Menu size={19} />
        </button>

        <div className="flex items-center gap-2">
          <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-gradient-to-br from-blue-600 to-indigo-600 text-white shadow-md">
            <Layers size={18} />
          </div>
          <div className="leading-none">
            <p className="text-base font-bold tracking-tight text-slate-900 dark:text-white">Bench</p>
            <p className="text-[11px] text-slate-400 dark:text-slate-500">Team Productivity Portal</p>
          </div>
        </div>

        <div className="ml-auto flex items-center gap-2">
          <div className="relative hidden sm:block">
            <Search size={15} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              value={search}
              onChange={(e) => onSearch(e.target.value)}
              placeholder="Search Bench…"
              className="w-52 rounded-xl border border-slate-200 bg-slate-50 py-2 pl-9 pr-3 text-sm text-slate-700 outline-none transition focus:w-64 focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200 dark:focus:bg-slate-950"
            />
          </div>

          <IconButton title="Notifications" className="relative">
            <Bell size={16} />
            <span className="absolute right-1.5 top-1.5 h-1.5 w-1.5 rounded-full bg-red-500" />
          </IconButton>

          <IconButton title="Toggle theme" onClick={onToggleTheme}>
            {theme === "dark" ? <Sun size={16} /> : <Moon size={16} />}
          </IconButton>

          <div className="relative">
            <button
              onClick={() => setMenuOpen((o) => !o)}
              className="ml-1 flex items-center gap-2 rounded-xl border border-slate-200 bg-white py-1 pl-1 pr-2.5 transition hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-900 dark:hover:bg-slate-800"
            >
              <div className="flex h-7 w-7 items-center justify-center rounded-lg bg-indigo-600 text-xs font-bold text-white">{initials}</div>
              <div className="hidden leading-none text-left sm:block">
                <p className="text-xs font-semibold text-slate-800 dark:text-slate-100">{user.name}</p>
                <p className="text-[10px] text-slate-400">{user.role}</p>
              </div>
              <ChevronDown size={13} className={cx("hidden shrink-0 text-slate-400 transition-transform sm:block", menuOpen && "rotate-180")} />
            </button>

            {menuOpen && (
              <>
                <div className="fixed inset-0 z-10" onClick={() => setMenuOpen(false)} />
                <div className="absolute right-0 top-12 z-20 w-60 animate-pop-in rounded-xl border border-slate-200 bg-white p-2 shadow-xl dark:border-slate-800 dark:bg-slate-900">
                  <div className="border-b border-slate-100 px-2 pb-2 dark:border-slate-800">
                    <p className="text-sm font-semibold text-slate-800 dark:text-slate-100">{user.name}</p>
                    <p className="truncate text-xs text-slate-400">{user.email}</p>
                    <div className="mt-1.5 inline-flex items-center gap-1.5 rounded-full bg-blue-50 px-2 py-0.5 text-[11px] font-medium text-blue-700 ring-1 ring-inset ring-blue-200 dark:bg-blue-500/10 dark:text-blue-300 dark:ring-blue-500/30">
                      {user.provider === "google" ? (
                        <GoogleIcon size={11} />
                      ) : user.provider === "microsoft" ? (
                        <MicrosoftIcon size={11} />
                      ) : user.provider === "guest" ? (
                        <CircleUserRound size={11} />
                      ) : (
                        <Mail size={11} />
                      )}
                      {user.provider === "google"
                        ? "Signed in with Google"
                        : user.provider === "microsoft"
                        ? "Signed in with Microsoft"
                        : user.provider === "guest"
                        ? "Guest session — not saved"
                        : "Signed in with Email"}
                    </div>
                    {user.provider === "microsoft" && accessLevel && (
                      <div className="mt-1.5 inline-flex items-center gap-1.5 rounded-full bg-purple-50 px-2 py-0.5 text-[11px] font-medium text-purple-700 ring-1 ring-inset ring-purple-200 dark:bg-purple-500/10 dark:text-purple-300 dark:ring-purple-500/30">
                        <ShieldCheck size={11} /> Access Level: {accessLevel}
                      </div>
                    )}
                  </div>
                  <button
                    onClick={() => {
                      setMenuOpen(false);
                      onLogout();
                    }}
                    className="mt-2 flex w-full items-center gap-2 rounded-lg px-2 py-1.5 text-sm font-medium text-red-600 transition hover:bg-red-50 dark:text-red-400 dark:hover:bg-red-500/10"
                  >
                    <LogOut size={14} /> Sign out
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    </header>
  );
}

function SidebarNav({ activeView, onNavigate, collapsed, items }) {
  return (
    <nav className="flex flex-1 flex-col gap-1 overflow-y-auto p-3">
      {items.map((item) => {
        const isActive = activeView === item.key;
        return (
          <button
            key={item.key}
            onClick={() => onNavigate(item.key)}
            title={collapsed ? item.label : undefined}
            className={cx(
              "flex items-center gap-3 rounded-xl px-3 py-2.5 text-sm font-medium transition",
              collapsed && "justify-center",
              isActive
                ? "bg-blue-600 text-white shadow-md shadow-blue-600/20"
                : "text-slate-600 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800"
            )}
          >
            <item.icon size={18} className="shrink-0" />
            {!collapsed && <span className="truncate">{item.label}</span>}
          </button>
        );
      })}
    </nav>
  );
}

function Sidebar({ activeView, onNavigate, collapsed, onToggleCollapse, items }) {
  return (
    <aside
      className={cx(
        "sticky top-16 hidden h-[calc(100vh-4rem)] shrink-0 flex-col border-r border-slate-200/80 bg-white/60 backdrop-blur-xl transition-all duration-300 dark:border-slate-800/80 dark:bg-slate-950/50 md:flex",
        collapsed ? "w-[76px]" : "w-64"
      )}
    >
      <SidebarNav activeView={activeView} onNavigate={onNavigate} collapsed={collapsed} items={items} />
      <div className="border-t border-slate-200/80 p-3 dark:border-slate-800/80">
        <button
          onClick={onToggleCollapse}
          className="flex w-full items-center justify-center gap-2 rounded-xl px-3 py-2 text-xs font-medium text-slate-500 transition hover:bg-slate-100 dark:text-slate-400 dark:hover:bg-slate-800"
        >
          {collapsed ? <PanelLeftOpen size={16} /> : <PanelLeftClose size={16} />}
          {!collapsed && "Collapse"}
        </button>
      </div>
    </aside>
  );
}

function MobileNav({ open, onClose, activeView, onNavigate, items }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-50 md:hidden">
      <div className="absolute inset-0 bg-slate-900/50 backdrop-blur-sm" onClick={onClose} />
      <div className="relative z-10 flex h-full w-72 flex-col border-r border-slate-200 bg-white shadow-2xl dark:border-slate-800 dark:bg-slate-950 animate-pop-in">
        <div className="flex h-16 items-center justify-between border-b border-slate-200 px-4 dark:border-slate-800">
          <div className="flex items-center gap-2">
            <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-gradient-to-br from-blue-600 to-indigo-600 text-white">
              <Layers size={16} />
            </div>
            <span className="font-bold text-slate-900 dark:text-white">Bench</span>
          </div>
          <button onClick={onClose} className="rounded-lg p-1 text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800">
            <X size={18} />
          </button>
        </div>
        <SidebarNav
          activeView={activeView}
          onNavigate={(key) => {
            onNavigate(key);
            onClose();
          }}
          collapsed={false}
          items={items}
        />
      </div>
    </div>
  );
}

/* ----------------------------------------------------------------------- */
/* Dashboard view                                                           */
/* ----------------------------------------------------------------------- */

function WeeklyChart({ workLogs }) {
  const order = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
  const totals = useMemo(() => {
    const map = Object.fromEntries(order.map((d) => [d, 0]));
    workLogs.forEach((l) => {
      const day = weekdayLabel(l.date);
      if (map[day] !== undefined) map[day] += Number(l.hours || 0);
    });
    return map;
  }, [workLogs]);
  const max = Math.max(1, ...Object.values(totals));

  return (
    <div className="flex h-48 items-end justify-between gap-3 px-1">
      {order.map((day) => {
        const value = totals[day];
        const pct = (value / max) * 100;
        return (
          <div key={day} className="flex flex-1 flex-col items-center gap-2">
            <span className="text-xs font-semibold text-slate-500 dark:text-slate-400">{value > 0 ? value.toFixed(1) : ""}</span>
            <div className="flex h-32 w-full items-end overflow-hidden rounded-lg bg-slate-100 dark:bg-slate-800">
              <div
                className="w-full rounded-lg bg-gradient-to-t from-blue-600 to-indigo-400 transition-all duration-700"
                style={{ height: `${value > 0 ? Math.max(6, pct) : 0}%` }}
              />
            </div>
            <span className="text-xs font-medium text-slate-400">{day}</span>
          </div>
        );
      })}
    </div>
  );
}

function KpiCard({ icon: Icon, label, value, tone, sublabel }) {
  const tones = {
    blue: "from-blue-600 to-indigo-600",
    green: "from-emerald-500 to-teal-500",
    orange: "from-orange-500 to-amber-500",
    purple: "from-purple-500 to-fuchsia-500",
  };
  return (
    <Card className="p-5">
      <div className="flex items-start justify-between">
        <div>
          <p className="text-xs font-medium uppercase tracking-wide text-slate-400">{label}</p>
          <p className="mt-2 text-2xl font-bold text-slate-900 dark:text-white">{value}</p>
          {sublabel && <p className="mt-1 text-xs text-emerald-600 dark:text-emerald-400">{sublabel}</p>}
        </div>
        <div className={cx("flex h-11 w-11 items-center justify-center rounded-xl bg-gradient-to-br text-white shadow-md", tones[tone] || tones.blue)}>
          <Icon size={19} />
        </div>
      </div>
    </Card>
  );
}

function DashboardView({ workLogs, challenges, learningItems, onNavigate, user }) {
  const totalHours = workLogs.reduce((s, l) => s + Number(l.hours || 0), 0);
  const completed = workLogs.filter((l) => l.status === "Completed").length;
  const productivity = Math.round((completed / Math.max(1, workLogs.length)) * 100);
  const resolvedChallenges = challenges.filter((c) => c.status === "Resolved").length;
  const activeCerts = learningItems.filter((l) => l.type === "Certification" && l.status === "Completed").length;

  const announcementIcon = { security: ShieldAlert, info: Bell, success: CheckCircle2 };
  const announcementIconWrap = {
    security: "bg-red-50 dark:bg-red-500/10",
    info: "bg-blue-50 dark:bg-blue-500/10",
    success: "bg-emerald-50 dark:bg-emerald-500/10",
  };

  return (
    <div className="space-y-6">
      <div className="overflow-hidden rounded-2xl bg-gradient-to-br from-blue-600 via-indigo-600 to-purple-600 p-6 text-white shadow-lg sm:p-8">
        <p className="text-sm font-medium text-blue-100">Welcome back,</p>
        <h1 className="mt-1 text-2xl font-bold sm:text-3xl">{user.name} 👋</h1>
        <p className="mt-2 max-w-xl text-sm text-blue-100">
          You've logged {totalHours.toFixed(1)}h this week across {new Set(workLogs.map((l) => l.project)).size} projects. {resolvedChallenges} technical
          challenges resolved. Keep it up!
        </p>
        <div className="mt-4 flex flex-wrap gap-2">
          <Badge tone="slate">
            <ShieldCheck size={12} /> Entra ID: All systems normal
          </Badge>
          <Badge tone="slate">
            <Flame size={12} /> {challenges.filter((c) => c.status !== "Resolved").length} open challenge(s)
          </Badge>
        </div>
      </div>

      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <KpiCard icon={Clock} label="Total Hours Logged" value={`${totalHours.toFixed(1)}h`} tone="blue" sublabel="This week" />
        <KpiCard icon={TrendingUp} label="Productivity Score" value={`${productivity}%`} tone="green" sublabel="+4% vs last week" />
        <KpiCard icon={CheckCircle2} label="Resolved Challenges" value={resolvedChallenges} tone="orange" sublabel={`of ${challenges.length} total`} />
        <KpiCard icon={Award} label="Active Certifications" value={activeCerts} tone="purple" sublabel="Microsoft Certified" />
      </div>

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-3">
        <Card className="p-5 lg:col-span-2">
          <div className="mb-4 flex items-center justify-between">
            <h3 className="flex items-center gap-2 font-semibold text-slate-800 dark:text-slate-100">
              <BarChart3 size={17} className="text-blue-600" /> Weekly Productivity
            </h3>
            <Badge tone="blue">Hours / day</Badge>
          </div>
          <WeeklyChart workLogs={workLogs} />
        </Card>

        <Card className="p-5">
          <h3 className="mb-4 flex items-center gap-2 font-semibold text-slate-800 dark:text-slate-100">
            <Rocket size={17} className="text-indigo-600" /> Quick Launchpad
          </h3>
          <div className="space-y-2">
            {QUICK_LAUNCH.map((q) => (
              <a
                key={q.id}
                href={q.href || "#"}
                target={q.href ? "_blank" : undefined}
                rel={q.href ? "noopener noreferrer" : undefined}
                onClick={(e) => {
                  if (!q.href) {
                    e.preventDefault();
                    onNavigate(q.view);
                  }
                }}
                className="flex items-center justify-between rounded-xl border border-slate-200 px-3 py-2.5 text-sm font-medium text-slate-700 transition hover:border-blue-300 hover:bg-blue-50 dark:border-slate-800 dark:text-slate-300 dark:hover:border-blue-500/40 dark:hover:bg-blue-500/10"
              >
                <span className="flex items-center gap-2">
                  <q.icon size={15} className="text-slate-400" /> {q.label}
                </span>
                <ExternalLink size={13} className="text-slate-300" />
              </a>
            ))}
          </div>
        </Card>
      </div>

      <Card className="p-5">
        <h3 className="mb-4 flex items-center gap-2 font-semibold text-slate-800 dark:text-slate-100">
          <Bell size={17} className="text-blue-600" /> Team Announcements &amp; Security Notices
        </h3>
        <div className="space-y-3">
          {ANNOUNCEMENTS.map((a) => {
            const Icon = announcementIcon[a.type];
            return (
              <div key={a.id} className="flex items-start gap-3 rounded-xl border border-slate-100 p-3 dark:border-slate-800">
                <div className={cx("mt-0.5 flex h-8 w-8 shrink-0 items-center justify-center rounded-lg", announcementIconWrap[a.type])}>
                  <Icon size={15} className={cx(a.type === "security" && "text-red-600 dark:text-red-400", a.type === "info" && "text-blue-600 dark:text-blue-400", a.type === "success" && "text-emerald-600 dark:text-emerald-400")} />
                </div>
                <div className="flex-1">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <p className="text-sm font-semibold text-slate-800 dark:text-slate-100">{a.title}</p>
                    <span className="text-xs text-slate-400">{formatDateShort(a.date)}</span>
                  </div>
                  <p className="mt-0.5 text-sm text-slate-500 dark:text-slate-400">{a.body}</p>
                </div>
              </div>
            );
          })}
        </div>
      </Card>
    </div>
  );
}

/* ----------------------------------------------------------------------- */
/* Work Log view                                                            */
/* ----------------------------------------------------------------------- */

const statusTone = { Completed: "green", "In Progress": "blue", Blocked: "red" };
const severityTone = { Critical: "red", High: "orange", Medium: "yellow" };
const priorityTone = { High: "red", Medium: "orange", Low: "slate" };

function WorkLogView({ workLogs, onAdd, onDelete, syncBadge, projectOptions }) {
  const options = projectOptions && projectOptions.length ? projectOptions : PROJECTS;
  const emptyForm = { taskTitle: "", project: options[0], hours: "", category: WORK_CATEGORIES[0], status: WORK_STATUSES[0], commit: "", jiraLink: "" };
  const [form, setForm] = useState(emptyForm);

  useEffect(() => {
    if (!options.includes(form.project)) {
      setForm((f) => ({ ...f, project: options[0] }));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [options]);

  const update = (field) => (e) => setForm((f) => ({ ...f, [field]: e.target.value }));

  const submit = (e) => {
    e.preventDefault();
    if (!form.taskTitle.trim() || !form.hours) return;
    onAdd({
      id: Date.now(),
      date: isoDaysAgo(0),
      taskTitle: form.taskTitle.trim(),
      project: form.project,
      hours: Number(form.hours),
      category: form.category,
      status: form.status,
      commit: form.commit.trim() || "-",
      jiraLink: form.jiraLink.trim(),
    });
    setForm({ ...emptyForm, project: options[0] });
  };

  const totalHours = workLogs.reduce((s, l) => s + Number(l.hours || 0), 0);

  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <div className="flex flex-wrap items-center gap-2">
            <h2 className="text-xl font-bold text-slate-900 dark:text-white">Daily Work Log &amp; Timesheets</h2>
            {syncBadge}
          </div>
          <p className="text-sm text-slate-500 dark:text-slate-400">{workLogs.length} entries · {totalHours.toFixed(1)}h logged</p>
        </div>
        <button
          onClick={() => downloadCSV(workLogs, "bench-work-logs.csv")}
          className="inline-flex items-center gap-2 rounded-xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-slate-700 dark:bg-white dark:text-slate-900 dark:hover:bg-slate-200"
        >
          <Download size={15} /> Export CSV
        </button>
      </div>

      <Card className="p-5">
        <h3 className="mb-4 flex items-center gap-2 font-semibold text-slate-800 dark:text-slate-100">
          <Plus size={16} className="text-blue-600" /> Log New Work Item
        </h3>
        <form onSubmit={submit} className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
          <div className="lg:col-span-2">
            <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Task Title</label>
            <input
              value={form.taskTitle}
              onChange={update("taskTitle")}
              required
              placeholder="e.g. Fix Graph subscription renewal bug"
              className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
            />
          </div>
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Project</label>
            <select
              value={form.project}
              onChange={update("project")}
              className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
            >
              {options.map((p) => (
                <option key={p}>{p}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Hours Spent</label>
            <input
              type="number"
              min="0"
              step="0.25"
              value={form.hours}
              onChange={update("hours")}
              required
              placeholder="2.5"
              className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
            />
          </div>
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Category</label>
            <select
              value={form.category}
              onChange={update("category")}
              className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
            >
              {WORK_CATEGORIES.map((c) => (
                <option key={c}>{c}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Status</label>
            <select
              value={form.status}
              onChange={update("status")}
              className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
            >
              {WORK_STATUSES.map((s) => (
                <option key={s}>{s}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Git Commit Signature (optional)</label>
            <input
              value={form.commit}
              onChange={update("commit")}
              placeholder="a3f9e21"
              className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 font-mono text-sm outline-none focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
            />
          </div>
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Jira Ticket Link (optional)</label>
            <input
              type="url"
              value={form.jiraLink}
              onChange={update("jiraLink")}
              placeholder="https://yourteam.atlassian.net/browse/PROJ-123"
              className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
            />
          </div>
          <div className="flex items-end">
            <button type="submit" className="inline-flex w-full items-center justify-center gap-2 rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-700">
              <Plus size={15} /> Add Entry
            </button>
          </div>
        </form>
      </Card>

      <Card className="overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full min-w-[840px] text-left text-sm">
            <thead className="border-b border-slate-200 bg-slate-50 text-xs uppercase tracking-wide text-slate-500 dark:border-slate-800 dark:bg-slate-900/60 dark:text-slate-400">
              <tr>
                <th className="px-4 py-3 font-medium">Task</th>
                <th className="px-4 py-3 font-medium">Project</th>
                <th className="px-4 py-3 font-medium">Hours</th>
                <th className="px-4 py-3 font-medium">Category</th>
                <th className="px-4 py-3 font-medium">Status</th>
                <th className="px-4 py-3 font-medium">Commit</th>
                <th className="px-4 py-3 font-medium">Jira</th>
                <th className="px-4 py-3 font-medium">Date</th>
                <th className="px-4 py-3" />
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
              {workLogs.map((log) => (
                <tr key={log.id} className="transition hover:bg-slate-50 dark:hover:bg-slate-800/40">
                  <td className="px-4 py-3 font-medium text-slate-800 dark:text-slate-100">{log.taskTitle}</td>
                  <td className="px-4 py-3 text-slate-500 dark:text-slate-400">{log.project}</td>
                  <td className="px-4 py-3 font-semibold text-slate-700 dark:text-slate-200">{Number(log.hours).toFixed(2)}</td>
                  <td className="px-4 py-3">
                    <Badge tone="purple">{log.category}</Badge>
                  </td>
                  <td className="px-4 py-3">
                    <Badge tone={statusTone[log.status]}>{log.status}</Badge>
                  </td>
                  <td className="px-4 py-3 font-mono text-xs text-slate-400">{log.commit}</td>
                  <td className="px-4 py-3">
                    {log.jiraLink ? (
                      <a
                        href={log.jiraLink}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="inline-flex items-center gap-1 text-blue-600 hover:underline dark:text-blue-400"
                        title={log.jiraLink}
                      >
                        <Link2 size={12} /> Ticket
                      </a>
                    ) : (
                      <span className="text-slate-300 dark:text-slate-600">—</span>
                    )}
                  </td>
                  <td className="px-4 py-3 text-slate-400">{formatDateShort(log.date)}</td>
                  <td className="px-4 py-3 text-right">
                    <button onClick={() => onDelete(log.id)} className="rounded-lg p-1.5 text-slate-300 transition hover:bg-red-50 hover:text-red-500 dark:hover:bg-red-500/10">
                      <Trash2 size={14} />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {workLogs.length === 0 && <EmptyState icon={ClipboardList} message="No work log entries yet. Add your first one above." />}
      </Card>
    </div>
  );
}

/* ----------------------------------------------------------------------- */
/* Projects view                                                           */
/* ----------------------------------------------------------------------- */

const fieldClass =
  "w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900";

function FormField({ label, className = "", children }) {
  return (
    <div className={className}>
      <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">{label}</label>
      {children}
    </div>
  );
}

function FormSection({ icon: Icon, title, first = false, children }) {
  return (
    <div className={cx("space-y-3", !first && "mt-5 border-t border-slate-100 pt-5 dark:border-slate-800")}>
      <div className="flex items-center gap-2">
        <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-blue-50 text-blue-600 dark:bg-blue-500/10 dark:text-blue-400">
          <Icon size={13} />
        </span>
        <h4 className="text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">{title}</h4>
      </div>
      {children}
    </div>
  );
}

// Shared by SearchableSelect / SearchableMultiSelect: renders the option
// panel through a portal into document.body, positioned by the trigger's own
// bounding rect. A plain `absolute` panel would get silently clipped by the
// Modal's `overflow-y-auto` content area whenever the field sits near the
// bottom of the scrollable region — portaling escapes that entirely, the
// same way a native <select>'s options escape page layout.
function DropdownPanel({ anchorRef, onClose, children }) {
  const panelRef = useRef(null);
  const [rect, setRect] = useState(null);

  useEffect(() => {
    const update = () => {
      const r = anchorRef.current?.getBoundingClientRect();
      if (r) setRect({ top: r.bottom + 4, left: r.left, width: r.width });
    };
    update();
    window.addEventListener("resize", update);
    // Scroll happens on the Modal's inner content area (or the page itself
    // for non-modal usage) — listening on the capture phase catches both
    // without needing a ref to that ancestor.
    window.addEventListener("scroll", update, true);
    return () => {
      window.removeEventListener("resize", update);
      window.removeEventListener("scroll", update, true);
    };
  }, [anchorRef]);

  useEffect(() => {
    const handleClick = (e) => {
      if (anchorRef.current?.contains(e.target)) return;
      if (panelRef.current?.contains(e.target)) return;
      onClose();
    };
    // Modal (the usual host for this panel) also closes on Escape via its
    // own window keydown listener. Registering this one on the capture phase
    // and stopping propagation lets Escape close just this panel first,
    // instead of the keydown reaching Modal's bubble-phase listener too and
    // closing the whole thing in the same keypress.
    const handleKey = (e) => {
      if (e.key !== "Escape") return;
      e.stopPropagation();
      onClose();
    };
    document.addEventListener("mousedown", handleClick);
    window.addEventListener("keydown", handleKey, true);
    return () => {
      document.removeEventListener("mousedown", handleClick);
      window.removeEventListener("keydown", handleKey, true);
    };
  }, [anchorRef, onClose]);

  if (!rect) return null;
  return createPortal(
    <div
      ref={panelRef}
      style={{ position: "fixed", top: rect.top, left: rect.left, width: rect.width }}
      className="z-[70] max-h-56 overflow-y-auto rounded-xl border border-slate-200 bg-white p-1 shadow-2xl dark:border-slate-700 dark:bg-slate-800"
    >
      {children}
    </div>,
    document.body
  );
}

function normalizeOptions(options) {
  return options.map((o) => (typeof o === "string" ? { value: o, label: o } : o));
}

function SearchableSelect({ value, onChange, options, placeholder = "Select…", allowEmpty = false, emptyLabel = "Unassigned" }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const triggerRef = useRef(null);
  const searchRef = useRef(null);

  const allOptions = allowEmpty ? [{ value: "", label: emptyLabel }, ...normalizeOptions(options)] : normalizeOptions(options);
  const filtered = query ? allOptions.filter((o) => o.label.toLowerCase().includes(query.toLowerCase())) : allOptions;
  const selected = allOptions.find((o) => o.value === value);

  const openPanel = () => {
    setQuery("");
    setOpen(true);
    setTimeout(() => searchRef.current?.focus(), 0);
  };

  return (
    <>
      <button
        type="button"
        ref={triggerRef}
        onClick={() => (open ? setOpen(false) : openPanel())}
        className={cx(fieldClass, "flex items-center justify-between gap-2 text-left")}
      >
        <span className={cx("truncate", !selected?.label && !selected?.value && !value && "text-slate-400")}>{selected ? selected.label || emptyLabel : placeholder}</span>
        <ChevronDown size={14} className={cx("shrink-0 text-slate-400 transition-transform", open && "rotate-180")} />
      </button>
      {open && (
        <DropdownPanel anchorRef={triggerRef} onClose={() => setOpen(false)}>
          <div className="sticky top-0 border-b border-slate-100 bg-white p-1.5 dark:border-slate-700 dark:bg-slate-800">
            <input
              ref={searchRef}
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search…"
              className="w-full rounded-lg border border-slate-200 bg-slate-50 px-2 py-1.5 text-sm outline-none focus:border-blue-400 dark:border-slate-600 dark:bg-slate-900"
            />
          </div>
          <div className="p-1">
            {filtered.length === 0 && <div className="px-3 py-2 text-sm text-slate-400">No matches</div>}
            {filtered.map((o) => (
              <button
                key={o.value || "__empty"}
                type="button"
                onClick={() => {
                  onChange(o.value);
                  setOpen(false);
                }}
                className={cx(
                  "flex w-full items-center justify-between rounded-lg px-3 py-1.5 text-left text-sm transition hover:bg-slate-100 dark:hover:bg-slate-700",
                  o.value === value ? "text-blue-700 dark:text-blue-300" : "text-slate-700 dark:text-slate-200"
                )}
              >
                {o.label}
                {o.value === value && <Check size={14} />}
              </button>
            ))}
          </div>
        </DropdownPanel>
      )}
    </>
  );
}

function SearchableMultiSelect({ value, onChange, options, placeholder = "Select…" }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const triggerRef = useRef(null);
  const searchRef = useRef(null);

  const allOptions = normalizeOptions(options);
  const filtered = query ? allOptions.filter((o) => o.label.toLowerCase().includes(query.toLowerCase())) : allOptions;
  const selectedSet = new Set(value);

  const openPanel = () => {
    setQuery("");
    setOpen(true);
    setTimeout(() => searchRef.current?.focus(), 0);
  };

  const toggle = (v) => {
    onChange(selectedSet.has(v) ? value.filter((x) => x !== v) : [...value, v]);
  };

  return (
    <>
      <button
        type="button"
        ref={triggerRef}
        onClick={() => (open ? setOpen(false) : openPanel())}
        className={cx(fieldClass, "flex min-h-[38px] items-center justify-between gap-2 text-left")}
      >
        {value.length ? (
          <span className="flex flex-1 flex-wrap gap-1">
            {value.map((v) => (
              <span
                key={v}
                className="inline-flex items-center gap-1 rounded-md bg-blue-50 px-1.5 py-0.5 text-xs font-medium text-blue-700 dark:bg-blue-500/10 dark:text-blue-300"
              >
                {v}
                <X
                  size={11}
                  className="cursor-pointer"
                  onClick={(e) => {
                    e.stopPropagation();
                    toggle(v);
                  }}
                />
              </span>
            ))}
          </span>
        ) : (
          <span className="text-slate-400">{placeholder}</span>
        )}
        <ChevronDown size={14} className={cx("shrink-0 text-slate-400 transition-transform", open && "rotate-180")} />
      </button>
      {open && (
        <DropdownPanel anchorRef={triggerRef} onClose={() => setOpen(false)}>
          <div className="sticky top-0 border-b border-slate-100 bg-white p-1.5 dark:border-slate-700 dark:bg-slate-800">
            <input
              ref={searchRef}
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search…"
              className="w-full rounded-lg border border-slate-200 bg-slate-50 px-2 py-1.5 text-sm outline-none focus:border-blue-400 dark:border-slate-600 dark:bg-slate-900"
            />
          </div>
          <div className="p-1">
            {filtered.length === 0 && <div className="px-3 py-2 text-sm text-slate-400">No matches</div>}
            {filtered.map((o) => (
              <button
                key={o.value}
                type="button"
                onClick={() => toggle(o.value)}
                className={cx(
                  "flex w-full items-center justify-between rounded-lg px-3 py-1.5 text-left text-sm transition hover:bg-slate-100 dark:hover:bg-slate-700",
                  selectedSet.has(o.value) ? "text-blue-700 dark:text-blue-300" : "text-slate-700 dark:text-slate-200"
                )}
              >
                {o.label}
                {selectedSet.has(o.value) && <Check size={14} />}
              </button>
            ))}
          </div>
        </DropdownPanel>
      )}
    </>
  );
}

function projectStatusTone(status) {
  if (status === "Active") return "blue";
  if (status === "Completed") return "green";
  if (status === "On Hold") return "yellow";
  if (status === "Cancelled") return "red";
  return "slate";
}

function healthTone(health) {
  if (health === "Green") return "green";
  if (health === "Yellow") return "yellow";
  if (health === "Red") return "red";
  return "slate";
}

function DetailRow({ label, value }) {
  return (
    <div>
      <div className="text-xs font-medium uppercase tracking-wide text-slate-400 dark:text-slate-500">{label}</div>
      <div className="mt-0.5 text-sm text-slate-800 dark:text-slate-100">{value || <span className="text-slate-400">—</span>}</div>
    </div>
  );
}

function ProjectsView({ projects, onAdd, onEdit, onDelete, syncBadge, peopleOptions }) {
  const people = peopleOptions && peopleOptions.length ? peopleOptions : ASSIGNEES;
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [detail, setDetail] = useState(null);
  const [form, setForm] = useState(EMPTY_PROJECT_FORM);

  const update = (field) => (e) => setForm((f) => ({ ...f, [field]: e.target.value }));
  const setField = (field) => (val) => setForm((f) => ({ ...f, [field]: val }));

  const openAddForm = () => {
    setEditingId(null);
    setForm(EMPTY_PROJECT_FORM);
    setShowForm(true);
  };

  const openEditForm = (project) => {
    setDetail(null);
    setEditingId(project.id);
    setForm(projectToForm(project));
    setShowForm(true);
  };

  const submit = (e) => {
    e.preventDefault();
    if (!form.title.trim()) return;
    const payload = {
      ...form,
      title: form.title.trim(),
      code: form.code.trim(),
      client: form.client.trim(),
      zone: form.zone.trim(),
      estimatedHours: Number(form.estimatedHours) || 0,
    };
    if (editingId) onEdit(editingId, payload);
    else onAdd(payload);
    setForm(EMPTY_PROJECT_FORM);
    setEditingId(null);
    setShowForm(false);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <div className="flex flex-wrap items-center gap-2">
            <h2 className="text-xl font-bold text-slate-900 dark:text-white">Projects</h2>
            {syncBadge}
          </div>
          <p className="text-sm text-slate-500 dark:text-slate-400">
            {projects.length} {projects.length === 1 ? "project" : "projects"} · used by Work Log &amp; Timesheets' Project field
          </p>
        </div>
        <button
          onClick={openAddForm}
          className="inline-flex items-center gap-2 rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-700"
        >
          <Plus size={15} /> Add Project
        </button>
      </div>

      <Card className="overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full min-w-[760px] text-left text-sm">
            <thead className="border-b border-slate-200 bg-slate-50 text-xs uppercase tracking-wide text-slate-500 dark:border-slate-800 dark:bg-slate-900/60 dark:text-slate-400">
              <tr>
                <th className="px-4 py-3 font-medium">Project</th>
                <th className="px-4 py-3 font-medium">Status</th>
                <th className="px-4 py-3 font-medium">Health</th>
                <th className="px-4 py-3 font-medium">Category</th>
                <th className="px-4 py-3 font-medium">Manager</th>
                <th className="px-4 py-3 font-medium">Target End</th>
                <th className="px-4 py-3" />
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
              {projects.map((p) => (
                <tr key={p.id} className="cursor-pointer transition hover:bg-slate-50 dark:hover:bg-slate-800/40" onClick={() => setDetail(p)}>
                  <td className="px-4 py-3">
                    <div className="font-medium text-slate-800 dark:text-slate-100">{p.title}</div>
                    {p.code && <div className="text-xs text-slate-400">{p.code}</div>}
                  </td>
                  <td className="px-4 py-3">{p.status && <Badge tone={projectStatusTone(p.status)}>{p.status}</Badge>}</td>
                  <td className="px-4 py-3">{p.health && <Badge tone={healthTone(p.health)}>{p.health}</Badge>}</td>
                  <td className="px-4 py-3 text-slate-600 dark:text-slate-300">{p.category || "—"}</td>
                  <td className="px-4 py-3 text-slate-600 dark:text-slate-300">{p.projectManager || "—"}</td>
                  <td className="px-4 py-3 text-slate-600 dark:text-slate-300">{p.targetEndDate ? formatDateShort(p.targetEndDate) : "—"}</td>
                  <td className="px-4 py-3 text-right">
                    <div className="flex items-center justify-end gap-1">
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          openEditForm(p);
                        }}
                        className="rounded-lg p-1.5 text-slate-300 transition hover:bg-blue-50 hover:text-blue-600 dark:hover:bg-blue-500/10"
                        title="Edit project"
                      >
                        <Pencil size={14} />
                      </button>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          onDelete(p.id);
                        }}
                        className="rounded-lg p-1.5 text-slate-300 transition hover:bg-red-50 hover:text-red-500 dark:hover:bg-red-500/10"
                        title="Delete project"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {projects.length === 0 && <EmptyState icon={Briefcase} message="No projects yet. Add your first one above." />}
      </Card>

      <Modal
        open={showForm}
        onClose={() => {
          setShowForm(false);
          setEditingId(null);
        }}
        title={editingId ? "Edit Project" : "Add Project"}
        wide
        footer={
          <div className="flex items-center justify-end gap-2">
            <button
              type="button"
              onClick={() => {
                setShowForm(false);
                setEditingId(null);
              }}
              className="rounded-xl px-4 py-2 text-sm font-semibold text-slate-600 transition hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800"
            >
              Cancel
            </button>
            <button
              type="submit"
              form="add-project-form"
              className="inline-flex items-center gap-2 rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-700"
            >
              {editingId ? (
                <>
                  <Check size={15} /> Save Changes
                </>
              ) : (
                <>
                  <Plus size={15} /> Add Project
                </>
              )}
            </button>
          </div>
        }
      >
        <form id="add-project-form" onSubmit={submit}>
          <FormSection icon={Info} title="Basic Info" first>
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
              <FormField label="Project Name" className="sm:col-span-2">
                <input required value={form.title} onChange={update("title")} placeholder="e.g. Contoso Intranet Hub" className={fieldClass} />
              </FormField>
              <FormField label="Project Code">
                <input value={form.code} onChange={update("code")} placeholder="e.g. CTH-001" className={fieldClass} />
              </FormField>
            </div>
            <div className="mt-3 grid grid-cols-1 gap-3 sm:grid-cols-2">
              <FormField label="Client / Company">
                <input value={form.client} onChange={update("client")} className={fieldClass} />
              </FormField>
              <FormField label="Zone">
                <input value={form.zone} onChange={update("zone")} placeholder="e.g. NA-East" className={fieldClass} />
              </FormField>
            </div>
            <FormField label="Description" className="mt-3">
              <textarea rows={2} value={form.description} onChange={update("description")} className={fieldClass} />
            </FormField>
          </FormSection>

          <FormSection icon={Tag} title="Classification">
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
              <FormField label="Status">
                <SearchableSelect value={form.status} onChange={setField("status")} options={PROJECT_STATUSES} />
              </FormField>
              <FormField label="Project Health">
                <SearchableSelect value={form.health} onChange={setField("health")} options={PROJECT_HEALTHS} />
              </FormField>
              <FormField label="Category">
                <SearchableSelect value={form.category} onChange={setField("category")} options={PROJECT_CATEGORIES} />
              </FormField>
            </div>
          </FormSection>

          <FormSection icon={Users} title="People">
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
              <FormField label="Project Manager">
                <SearchableSelect value={form.projectManager} onChange={setField("projectManager")} options={people} allowEmpty />
              </FormField>
              <FormField label="Technical Lead">
                <SearchableSelect value={form.technicalLead} onChange={setField("technicalLead")} options={people} allowEmpty />
              </FormField>
              <FormField label="Team Members">
                <SearchableMultiSelect value={form.teamMembers} onChange={setField("teamMembers")} options={people} placeholder="Select team members…" />
              </FormField>
            </div>
          </FormSection>

          <FormSection icon={Clock} title="Timeline & Effort">
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
              <FormField label="Start Date">
                <input type="date" value={form.startDate} onChange={update("startDate")} className={fieldClass} />
              </FormField>
              <FormField label="Target End Date">
                <input type="date" value={form.targetEndDate} onChange={update("targetEndDate")} className={fieldClass} />
              </FormField>
              <FormField label="Actual End Date">
                <input type="date" value={form.actualEndDate} onChange={update("actualEndDate")} className={fieldClass} />
              </FormField>
            </div>
            <FormField label="Estimated Hours" className="mt-3 sm:w-1/3">
              <input type="number" min="0" step="1" value={form.estimatedHours} onChange={update("estimatedHours")} className={fieldClass} />
            </FormField>
          </FormSection>

          <FormSection icon={Link2} title="Links">
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
              <FormField label="Repository URL">
                <input type="url" value={form.repositoryUrl} onChange={update("repositoryUrl")} placeholder="https://github.com/…" className={fieldClass} />
              </FormField>
              <FormField label="Documentation Link">
                <input type="url" value={form.documentationLink} onChange={update("documentationLink")} placeholder="https://…" className={fieldClass} />
              </FormField>
            </div>
          </FormSection>
        </form>
      </Modal>

      <Modal
        open={Boolean(detail)}
        onClose={() => setDetail(null)}
        title={detail?.title || "Project"}
        wide
        footer={
          detail && (
            <div className="flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setDetail(null)}
                className="rounded-xl px-4 py-2 text-sm font-semibold text-slate-600 transition hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800"
              >
                Close
              </button>
              <button
                type="button"
                onClick={() => openEditForm(detail)}
                className="inline-flex items-center gap-2 rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-700"
              >
                <Pencil size={14} /> Edit Project
              </button>
            </div>
          )
        }
      >
        {detail && (
          <div className="space-y-5">
            <div className="flex flex-wrap gap-2">
              {detail.status && <Badge tone={projectStatusTone(detail.status)}>{detail.status}</Badge>}
              {detail.health && <Badge tone={healthTone(detail.health)}>Health: {detail.health}</Badge>}
              {detail.category && <Badge>{detail.category}</Badge>}
            </div>
            {detail.description && <p className="text-sm text-slate-600 dark:text-slate-300">{detail.description}</p>}
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
              <DetailRow label="Project Code" value={detail.code} />
              <DetailRow label="Client / Company" value={detail.client} />
              <DetailRow label="Zone" value={detail.zone} />
              <DetailRow label="Project Manager" value={detail.projectManager} />
              <DetailRow label="Technical Lead" value={detail.technicalLead} />
              <DetailRow label="Team Members" value={detail.teamMembers?.join(", ")} />
              <DetailRow label="Start Date" value={detail.startDate && formatDate(detail.startDate)} />
              <DetailRow label="Target End Date" value={detail.targetEndDate && formatDate(detail.targetEndDate)} />
              <DetailRow label="Actual End Date" value={detail.actualEndDate && formatDate(detail.actualEndDate)} />
              <DetailRow label="Estimated Hours" value={detail.estimatedHours ? `${detail.estimatedHours}h` : ""} />
              <DetailRow
                label="Repository"
                value={
                  detail.repositoryUrl && (
                    <a href={detail.repositoryUrl} target="_blank" rel="noopener noreferrer" className="inline-flex items-center gap-1 text-blue-600 hover:underline dark:text-blue-400">
                      {detail.repositoryUrl} <ExternalLink size={12} />
                    </a>
                  )
                }
              />
              <DetailRow
                label="Documentation"
                value={
                  detail.documentationLink && (
                    <a href={detail.documentationLink} target="_blank" rel="noopener noreferrer" className="inline-flex items-center gap-1 text-blue-600 hover:underline dark:text-blue-400">
                      {detail.documentationLink} <ExternalLink size={12} />
                    </a>
                  )
                }
              />
            </div>
          </div>
        )}
      </Modal>
    </div>
  );
}

/* ----------------------------------------------------------------------- */
/* Knowledge Base view                                                      */
/* ----------------------------------------------------------------------- */

function KnowledgeBaseView({ copiedKey, onCopy }) {
  const [query, setQuery] = useState("");
  const [category, setCategory] = useState("All");
  const [selected, setSelected] = useState(null);

  const categories = ["All", ...new Set(KB_ARTICLES.map((a) => a.category))];

  const filtered = KB_ARTICLES.filter((a) => {
    const matchesCategory = category === "All" || a.category === category;
    const q = query.toLowerCase();
    const matchesQuery = !q || a.title.toLowerCase().includes(q) || a.tags.some((t) => t.toLowerCase().includes(q));
    return matchesCategory && matchesQuery;
  });

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-xl font-bold text-slate-900 dark:text-white">Knowledge Base Repository</h2>
        <p className="text-sm text-slate-500 dark:text-slate-400">Searchable engineering knowledge — problems, root causes, and battle-tested solutions.</p>
      </div>

      <div className="flex flex-col gap-3 sm:flex-row sm:items-center">
        <div className="relative flex-1">
          <Search size={15} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search articles, tags, technologies…"
            className="w-full rounded-xl border border-slate-200 bg-white py-2.5 pl-9 pr-3 text-sm outline-none focus:border-blue-400 dark:border-slate-700 dark:bg-slate-900"
          />
        </div>
        <div className="flex flex-wrap gap-2">
          {categories.map((c) => (
            <button
              key={c}
              onClick={() => setCategory(c)}
              className={cx(
                "rounded-full px-3 py-1.5 text-xs font-semibold transition",
                category === c
                  ? "bg-blue-600 text-white shadow-sm"
                  : "border border-slate-200 bg-white text-slate-600 hover:border-blue-300 hover:text-blue-700 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-300"
              )}
            >
              {c}
            </button>
          ))}
        </div>
      </div>

      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
        {filtered.map((a) => (
          <Card key={a.id} className="flex flex-col p-5">
            <div className="mb-2 flex items-center justify-between">
              <Badge tone="blue">{a.category}</Badge>
              <span className="text-xs text-slate-400">{formatDateShort(a.updatedAt)}</span>
            </div>
            <h3 className="font-semibold text-slate-800 dark:text-slate-100">{a.title}</h3>
            <p className="mt-1.5 line-clamp-2 text-sm text-slate-500 dark:text-slate-400">{a.problem}</p>
            <div className="mt-3 flex flex-wrap gap-1.5">
              {a.tags.map((t) => (
                <span key={t} className="rounded-md bg-slate-100 px-2 py-0.5 text-[11px] font-medium text-slate-500 dark:bg-slate-800 dark:text-slate-400">
                  {t}
                </span>
              ))}
            </div>
            <button
              onClick={() => setSelected(a)}
              className="mt-4 inline-flex items-center gap-1.5 self-start text-sm font-semibold text-blue-600 transition hover:text-blue-700 dark:text-blue-400"
            >
              <BookMarked size={14} /> Read full article <ChevronRight size={14} />
            </button>
          </Card>
        ))}
        {filtered.length === 0 && (
          <div className="sm:col-span-2 lg:col-span-3">
            <EmptyState icon={Search} message="No articles match your search." />
          </div>
        )}
      </div>

      <Modal open={!!selected} onClose={() => setSelected(null)} title={selected?.title} wide>
        {selected && (
          <div className="space-y-5">
            <div className="flex flex-wrap items-center gap-2">
              <Badge tone="blue">{selected.category}</Badge>
              {selected.tags.map((t) => (
                <Badge key={t} tone="slate">
                  {t}
                </Badge>
              ))}
              <span className="text-xs text-slate-400">Updated {formatDate(selected.updatedAt)} by {selected.author}</span>
            </div>
            <div>
              <h4 className="mb-1 text-xs font-bold uppercase tracking-wide text-slate-400">Problem</h4>
              <p className="text-sm leading-relaxed text-slate-700 dark:text-slate-300">{selected.problem}</p>
            </div>
            <div>
              <h4 className="mb-1 text-xs font-bold uppercase tracking-wide text-slate-400">Root Cause</h4>
              <p className="text-sm leading-relaxed text-slate-700 dark:text-slate-300">{selected.rootCause}</p>
            </div>
            <div>
              <h4 className="mb-1 text-xs font-bold uppercase tracking-wide text-slate-400">Solution</h4>
              <p className="text-sm leading-relaxed text-slate-700 dark:text-slate-300">{selected.solution}</p>
            </div>
            <div>
              <h4 className="mb-2 text-xs font-bold uppercase tracking-wide text-slate-400">Code Snippet</h4>
              <CodeBlock code={selected.code} language={selected.language} copyKey={selected.id} copiedKey={copiedKey} onCopy={onCopy} />
            </div>
          </div>
        )}
      </Modal>
    </div>
  );
}

/* ----------------------------------------------------------------------- */
/* Challenges view                                                          */
/* ----------------------------------------------------------------------- */

function ChallengesView() {
  const [selected, setSelected] = useState(null);
  const [filter, setFilter] = useState("All");
  const severities = ["All", "Critical", "High", "Medium"];
  const filtered = CHALLENGES.filter((c) => filter === "All" || c.severity === filter);

  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 className="text-xl font-bold text-slate-900 dark:text-white">Technical Challenges Log</h2>
          <p className="text-sm text-slate-500 dark:text-slate-400">Incident tracking with root cause analysis and preventive follow-ups.</p>
        </div>
        <div className="flex gap-2">
          {severities.map((s) => (
            <button
              key={s}
              onClick={() => setFilter(s)}
              className={cx(
                "rounded-full px-3 py-1.5 text-xs font-semibold transition",
                filter === s ? "bg-slate-900 text-white dark:bg-white dark:text-slate-900" : "border border-slate-200 bg-white text-slate-600 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-300"
              )}
            >
              {s}
            </button>
          ))}
        </div>
      </div>

      <div className="space-y-3">
        {filtered.map((c) => (
          <Card key={c.id} className="cursor-pointer p-5 transition hover:border-blue-300 dark:hover:border-blue-500/40" onClick={() => setSelected(c)}>
            <div className="flex flex-wrap items-start justify-between gap-3">
              <div className="flex items-start gap-3">
                <div className={cx("mt-0.5 flex h-9 w-9 shrink-0 items-center justify-center rounded-xl", c.severity === "Critical" ? "bg-red-50 dark:bg-red-500/10" : c.severity === "High" ? "bg-orange-50 dark:bg-orange-500/10" : "bg-yellow-50 dark:bg-yellow-500/10")}>
                  <AlertTriangle size={16} className={cx(c.severity === "Critical" && "text-red-600", c.severity === "High" && "text-orange-600", c.severity === "Medium" && "text-yellow-600")} />
                </div>
                <div>
                  <p className="font-semibold text-slate-800 dark:text-slate-100">{c.title}</p>
                  <p className="mt-0.5 text-xs text-slate-400">{c.system} · {formatDate(c.date)}</p>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <Badge tone={severityTone[c.severity]}>{c.severity}</Badge>
                <Badge tone={c.status === "Resolved" ? "green" : "blue"}>{c.status}</Badge>
              </div>
            </div>
          </Card>
        ))}
      </div>

      <Modal open={!!selected} onClose={() => setSelected(null)} title={selected?.title} wide>
        {selected && (
          <div className="space-y-5">
            <div className="flex flex-wrap items-center gap-2">
              <Badge tone={severityTone[selected.severity]}>{selected.severity}</Badge>
              <Badge tone={selected.status === "Resolved" ? "green" : "blue"}>{selected.status}</Badge>
              <span className="text-xs text-slate-400">{selected.system} · {formatDate(selected.date)}</span>
            </div>
            <div>
              <h4 className="mb-1 text-xs font-bold uppercase tracking-wide text-slate-400">Root Cause Analysis</h4>
              <p className="text-sm leading-relaxed text-slate-700 dark:text-slate-300">{selected.rootCause}</p>
            </div>
            <div>
              <h4 className="mb-1 text-xs font-bold uppercase tracking-wide text-slate-400">Resolution Steps</h4>
              <ol className="ml-4 list-decimal space-y-1.5 text-sm leading-relaxed text-slate-700 dark:text-slate-300">
                {selected.resolutionSteps.map((s, i) => (
                  <li key={i}>{s}</li>
                ))}
              </ol>
            </div>
            <div className="flex items-center gap-2 rounded-xl border border-slate-100 p-3 dark:border-slate-800">
              <Clock size={15} className="text-blue-600" />
              <span className="text-sm font-medium text-slate-700 dark:text-slate-300">Resolution Time: {selected.resolutionTime}</span>
            </div>
            <div>
              <h4 className="mb-1 text-xs font-bold uppercase tracking-wide text-slate-400">Preventive Recommendations</h4>
              <p className="text-sm leading-relaxed text-slate-700 dark:text-slate-300">{selected.preventive}</p>
            </div>
          </div>
        )}
      </Modal>
    </div>
  );
}

/* ----------------------------------------------------------------------- */
/* Learning Center view                                                     */
/* ----------------------------------------------------------------------- */

function LearningCenterView() {
  const [filter, setFilter] = useState("All");
  const types = ["All", "Certification", "Course", "Workshop"];
  const filtered = LEARNING_ITEMS.filter((l) => filter === "All" || l.type === filter);

  const typeTone = { Certification: "purple", Course: "blue", Workshop: "orange" };
  const statusTone2 = { Completed: "green", "In Progress": "blue", Planned: "slate" };

  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 className="text-xl font-bold text-slate-900 dark:text-white">Learning Center &amp; Certifications</h2>
          <p className="text-sm text-slate-500 dark:text-slate-400">Track Microsoft Certifications, Pluralsight courses, and Learn workshops.</p>
        </div>
        <div className="flex gap-2">
          {types.map((t) => (
            <button
              key={t}
              onClick={() => setFilter(t)}
              className={cx(
                "rounded-full px-3 py-1.5 text-xs font-semibold transition",
                filter === t ? "bg-blue-600 text-white" : "border border-slate-200 bg-white text-slate-600 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-300"
              )}
            >
              {t}
            </button>
          ))}
        </div>
      </div>

      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
        {filtered.map((l) => (
          <Card key={l.id} className="flex flex-col p-5">
            <div className="mb-2 flex items-center justify-between">
              <Badge tone={typeTone[l.type]}>{l.type}</Badge>
              <Badge tone={statusTone2[l.status]}>{l.status}</Badge>
            </div>
            <h3 className="font-semibold text-slate-800 dark:text-slate-100">{l.title}</h3>
            <p className="mt-0.5 text-xs text-slate-400">{l.provider}</p>

            <div className="mt-4">
              <div className="mb-1.5 flex items-center justify-between text-xs">
                <span className="font-medium text-slate-500 dark:text-slate-400">Progress</span>
                <span className="font-semibold text-slate-700 dark:text-slate-200">{l.progress}%</span>
              </div>
              <ProgressBar value={l.progress} tone={l.status === "Completed" ? "green" : "blue"} />
            </div>

            <div className="mt-3 flex items-center gap-1.5 text-xs text-slate-400">
              <Calendar size={12} /> Target: {formatDate(l.dueDate)}
            </div>

            <div className="mt-3 flex items-start gap-1.5 rounded-lg bg-slate-50 p-2.5 text-xs leading-relaxed text-slate-500 dark:bg-slate-800/60 dark:text-slate-400">
              <GraduationCap size={13} className="mt-0.5 shrink-0" /> {l.notes}
            </div>
          </Card>
        ))}
      </div>
    </div>
  );
}

/* ----------------------------------------------------------------------- */
/* Components Library view                                                  */
/* ----------------------------------------------------------------------- */

function ComponentsLibraryView({ copiedKey, onCopy }) {
  const [query, setQuery] = useState("");
  const [category, setCategory] = useState("All");
  const [expanded, setExpanded] = useState(null);

  const categories = ["All", ...new Set(COMPONENT_LIBRARY.map((c) => c.category))];

  const filtered = COMPONENT_LIBRARY.filter((c) => {
    const matchesCategory = category === "All" || c.category === category;
    const q = query.toLowerCase();
    const matchesQuery = !q || c.name.toLowerCase().includes(q) || c.tags.some((t) => t.toLowerCase().includes(q));
    return matchesCategory && matchesQuery;
  });

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-xl font-bold text-slate-900 dark:text-white">Reusable Code &amp; Components Library</h2>
        <p className="text-sm text-slate-500 dark:text-slate-400">SPFx utilities, Power Apps formulas, Power Automate expressions, and Graph API scripts.</p>
      </div>

      <div className="flex flex-col gap-3 sm:flex-row sm:items-center">
        <div className="relative flex-1">
          <Search size={15} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search components…"
            className="w-full rounded-xl border border-slate-200 bg-white py-2.5 pl-9 pr-3 text-sm outline-none focus:border-blue-400 dark:border-slate-700 dark:bg-slate-900"
          />
        </div>
        <div className="flex flex-wrap gap-2">
          {categories.map((c) => (
            <button
              key={c}
              onClick={() => setCategory(c)}
              className={cx(
                "rounded-full px-3 py-1.5 text-xs font-semibold transition",
                category === c ? "bg-blue-600 text-white" : "border border-slate-200 bg-white text-slate-600 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-300"
              )}
            >
              {c}
            </button>
          ))}
        </div>
      </div>

      <div className="space-y-4">
        {filtered.map((c) => {
          const isOpen = expanded === c.id;
          return (
            <Card key={c.id} className="p-5">
              <div className="flex flex-wrap items-start justify-between gap-3">
                <div className="flex items-start gap-3">
                  <div className="mt-0.5 flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-blue-50 text-blue-600 dark:bg-blue-500/10 dark:text-blue-400">
                    <FileCode2 size={16} />
                  </div>
                  <div>
                    <div className="flex flex-wrap items-center gap-2">
                      <p className="font-semibold text-slate-800 dark:text-slate-100">{c.name}</p>
                      <Badge tone="slate">{c.version}</Badge>
                    </div>
                    <p className="mt-0.5 text-sm text-slate-500 dark:text-slate-400">{c.description}</p>
                    <div className="mt-2 flex flex-wrap gap-1.5">
                      <Badge tone="blue">{c.category}</Badge>
                      {c.tags.map((t) => (
                        <span key={t} className="rounded-md bg-slate-100 px-2 py-0.5 text-[11px] font-medium text-slate-500 dark:bg-slate-800 dark:text-slate-400">
                          {t}
                        </span>
                      ))}
                    </div>
                  </div>
                </div>
                <button
                  onClick={() => setExpanded(isOpen ? null : c.id)}
                  className="inline-flex items-center gap-1 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-semibold text-slate-600 transition hover:border-blue-300 hover:text-blue-700 dark:border-slate-700 dark:text-slate-300"
                >
                  {isOpen ? "Hide code" : "View code"} <ChevronDown size={13} className={cx("transition-transform", isOpen && "rotate-180")} />
                </button>
              </div>
              {isOpen && (
                <div className="mt-4">
                  <CodeBlock code={c.code} language={c.language} copyKey={c.id} copiedKey={copiedKey} onCopy={onCopy} />
                </div>
              )}
            </Card>
          );
        })}
        {filtered.length === 0 && <EmptyState icon={Code2} message="No components match your search." />}
      </div>
    </div>
  );
}

/* ----------------------------------------------------------------------- */
/* Issue Tracker (Kanban) view                                              */
/* ----------------------------------------------------------------------- */

const KANBAN_COLUMNS = [
  { key: "todo", label: "To Do", tone: "border-slate-300 dark:border-slate-700" },
  { key: "inprogress", label: "In Progress", tone: "border-blue-300 dark:border-blue-500/40" },
  { key: "critical", label: "Critical / At Risk", tone: "border-red-300 dark:border-red-500/40" },
];

function TicketCard({ ticket, onDragStart }) {
  return (
    <div
      draggable
      onDragStart={(e) => onDragStart(e, ticket.id)}
      className="cursor-grab rounded-xl border border-slate-200 bg-white p-3.5 shadow-sm transition active:cursor-grabbing hover:shadow-md dark:border-slate-800 dark:bg-slate-900"
    >
      <div className="flex items-start justify-between gap-2">
        <p className="text-sm font-semibold text-slate-800 dark:text-slate-100">{ticket.title}</p>
        <MoreVertical size={14} className="mt-0.5 shrink-0 text-slate-300" />
      </div>
      <p className="mt-1 line-clamp-2 text-xs text-slate-500 dark:text-slate-400">{ticket.description}</p>
      <div className="mt-2 flex flex-wrap gap-1">
        {ticket.tags.map((t) => (
          <span key={t} className="rounded-md bg-slate-100 px-1.5 py-0.5 text-[10px] font-medium text-slate-500 dark:bg-slate-800 dark:text-slate-400">
            {t}
          </span>
        ))}
      </div>
      <div className="mt-3 flex items-center justify-between">
        <div className="flex items-center gap-1.5">
          <div className="flex h-6 w-6 items-center justify-center rounded-full bg-indigo-100 text-[10px] font-bold text-indigo-700 dark:bg-indigo-500/20 dark:text-indigo-300">
            {ticket.assignee.split(" ").map((n) => n[0]).join("")}
          </div>
          <span className="text-xs text-slate-500 dark:text-slate-400">{ticket.assignee}</span>
        </div>
        <Badge tone={priorityTone[ticket.priority]}>{ticket.priority}</Badge>
      </div>
      <div className="mt-2 flex items-center gap-1 text-[11px] text-slate-400">
        <Calendar size={11} /> Due {formatDateShort(ticket.dueDate)}
      </div>
    </div>
  );
}

function IssueTrackerView({ issues, onCreateIssue, onMoveIssue, syncBadge, assigneeOptions }) {
  const assignees = assigneeOptions && assigneeOptions.length ? assigneeOptions : ASSIGNEES;
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ title: "", description: "", assignee: assignees[0], priority: "Medium", dueDate: isoDaysAgo(-7), tags: "" });

  useEffect(() => {
    if (!assignees.includes(form.assignee)) {
      setForm((f) => ({ ...f, assignee: assignees[0] }));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [assignees]);

  const handleDragStart = (e, id) => e.dataTransfer.setData("text/plain", String(id));
  const handleDrop = (e, columnKey) => {
    e.preventDefault();
    const id = e.dataTransfer.getData("text/plain");
    const ticket = issues.find((t) => String(t.id) === id);
    if (ticket && ticket.column !== columnKey) onMoveIssue(ticket.id, columnKey);
  };

  const submit = (e) => {
    e.preventDefault();
    if (!form.title.trim()) return;
    onCreateIssue({
      title: form.title.trim(),
      description: form.description.trim(),
      assignee: form.assignee,
      priority: form.priority,
      dueDate: form.dueDate,
      tags: form.tags ? form.tags.split(",").map((t) => t.trim()).filter(Boolean) : [],
      column: "todo",
    });
    setForm({ title: "", description: "", assignee: assignees[0], priority: "Medium", dueDate: isoDaysAgo(-7), tags: "" });
    setShowForm(false);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <div className="flex flex-wrap items-center gap-2">
            <h2 className="text-xl font-bold text-slate-900 dark:text-white">Issues, Risks &amp; Bugs Tracker</h2>
            {syncBadge}
          </div>
          <p className="text-sm text-slate-500 dark:text-slate-400">Drag cards between columns to update status.</p>
        </div>
        <button
          onClick={() => setShowForm(true)}
          className="inline-flex items-center gap-2 rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-700"
        >
          <Plus size={15} /> New Ticket
        </button>
      </div>

      <div className="grid grid-cols-1 gap-4 md:grid-cols-3">
        {KANBAN_COLUMNS.map((col) => {
          const items = issues.filter((t) => t.column === col.key);
          return (
            <div
              key={col.key}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => handleDrop(e, col.key)}
              className={cx("flex flex-col rounded-2xl border-2 border-dashed bg-slate-50/60 p-3 dark:bg-slate-900/30", col.tone)}
            >
              <div className="mb-3 flex items-center justify-between px-1">
                <h3 className="text-sm font-bold text-slate-700 dark:text-slate-200">{col.label}</h3>
                <span className="rounded-full bg-white px-2 py-0.5 text-xs font-semibold text-slate-500 shadow-sm dark:bg-slate-800 dark:text-slate-400">{items.length}</span>
              </div>
              <div className="flex flex-1 flex-col gap-3 min-h-[120px]">
                {items.map((t) => (
                  <TicketCard key={t.id} ticket={t} onDragStart={handleDragStart} />
                ))}
                {items.length === 0 && <div className="rounded-xl border border-dashed border-slate-200 p-4 text-center text-xs text-slate-400 dark:border-slate-700">Drop tickets here</div>}
              </div>
            </div>
          );
        })}
      </div>

      <Modal open={showForm} onClose={() => setShowForm(false)} title="New Ticket">
        <form onSubmit={submit} className="space-y-3">
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Title</label>
            <input
              required
              value={form.title}
              onChange={(e) => setForm((f) => ({ ...f, title: e.target.value }))}
              className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
            />
          </div>
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Description</label>
            <textarea
              rows={3}
              value={form.description}
              onChange={(e) => setForm((f) => ({ ...f, description: e.target.value }))}
              className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
            />
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Assignee</label>
              <select
                value={form.assignee}
                onChange={(e) => setForm((f) => ({ ...f, assignee: e.target.value }))}
                className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
              >
                {assignees.map((a) => (
                  <option key={a}>{a}</option>
                ))}
              </select>
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Priority</label>
              <select
                value={form.priority}
                onChange={(e) => setForm((f) => ({ ...f, priority: e.target.value }))}
                className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
              >
                {["Low", "Medium", "High"].map((p) => (
                  <option key={p}>{p}</option>
                ))}
              </select>
            </div>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Due Date</label>
              <input
                type="date"
                value={form.dueDate}
                onChange={(e) => setForm((f) => ({ ...f, dueDate: e.target.value }))}
                className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
              />
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Tags (comma separated)</label>
              <input
                value={form.tags}
                onChange={(e) => setForm((f) => ({ ...f, tags: e.target.value }))}
                placeholder="Bug, Security"
                className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none focus:border-blue-400 focus:bg-white dark:border-slate-700 dark:bg-slate-800 dark:focus:bg-slate-900"
              />
            </div>
          </div>
          <button type="submit" className="w-full rounded-xl bg-blue-600 px-4 py-2.5 text-sm font-semibold text-white transition hover:bg-blue-700">
            Create Ticket
          </button>
        </form>
      </Modal>
    </div>
  );
}

/* ----------------------------------------------------------------------- */
/* Architecture & Schema Admin view                                         */
/* ----------------------------------------------------------------------- */

function ArchitectureAdminView({ copiedKey, onCopy, spStatus, spError }) {
  const [tab, setTab] = useState("schemas");
  const tabs = [
    { key: "schemas", label: "SharePoint List Schemas", icon: Database },
    { key: "tree", label: "SPFx Directory Tree", icon: FolderTree },
    { key: "flows", label: "Power Automate Flows", icon: Workflow },
    { key: "connect", label: "Connect SharePoint", icon: Cloud },
  ];

  const envRows = [
    { key: "VITE_MSAL_CLIENT_ID", value: import.meta.env.VITE_MSAL_CLIENT_ID, note: "App Registration (client) ID" },
    { key: "VITE_MSAL_TENANT_ID", value: import.meta.env.VITE_MSAL_TENANT_ID || "common", note: "Directory (tenant) ID, or 'common'" },
    { key: "VITE_SP_HOSTNAME", value: import.meta.env.VITE_SP_HOSTNAME, note: "e.g. contoso.sharepoint.com" },
    { key: "VITE_SP_SITE_PATH", value: import.meta.env.VITE_SP_SITE_PATH, note: "e.g. /sites/BenchPortal" },
    { key: "VITE_AUTH_API_URL", value: import.meta.env.VITE_AUTH_API_URL, note: "Optional — see step 5 below" },
  ];

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-xl font-bold text-slate-900 dark:text-white">Architecture &amp; Schema Configuration</h2>
        <p className="text-sm text-slate-500 dark:text-slate-400">Admin reference for Bench's SharePoint Online data model, SPFx scaffolding, and Power Automate flows.</p>
      </div>

      <div className="flex flex-wrap gap-2">
        {tabs.map((t) => (
          <button
            key={t.key}
            onClick={() => setTab(t.key)}
            className={cx(
              "inline-flex items-center gap-2 rounded-xl px-3.5 py-2 text-sm font-semibold transition",
              tab === t.key ? "bg-slate-900 text-white dark:bg-white dark:text-slate-900" : "border border-slate-200 bg-white text-slate-600 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-300"
            )}
          >
            <t.icon size={15} /> {t.label}
          </button>
        ))}
      </div>

      {tab === "schemas" && (
        <div className="space-y-4">
          {SP_SCHEMAS.map((schema) => (
            <Card key={schema.name} className="p-5">
              <div className="mb-3 flex items-center gap-2">
                <Server size={16} className="text-blue-600" />
                <h3 className="font-mono text-sm font-bold text-slate-800 dark:text-slate-100">{schema.name}</h3>
              </div>
              <p className="mb-3 text-sm text-slate-500 dark:text-slate-400">{schema.description}</p>
              <div className="overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-800">
                <table className="w-full min-w-[520px] text-left text-sm">
                  <thead className="bg-slate-50 text-xs uppercase tracking-wide text-slate-500 dark:bg-slate-900/60 dark:text-slate-400">
                    <tr>
                      <th className="px-3 py-2 font-medium">Field</th>
                      <th className="px-3 py-2 font-medium">Type</th>
                      <th className="px-3 py-2 font-medium">Description</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                    {schema.fields.map((f) => (
                      <tr key={f.name}>
                        <td className="px-3 py-2 font-mono text-xs font-semibold text-blue-700 dark:text-blue-400">{f.name}</td>
                        <td className="px-3 py-2 text-xs text-slate-500 dark:text-slate-400">{f.type}</td>
                        <td className="px-3 py-2 text-slate-600 dark:text-slate-300">{f.description}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Card>
          ))}
        </div>
      )}

      {tab === "tree" && (
        <Card className="p-5">
          <div className="mb-3 flex items-center gap-2">
            <FolderTree size={16} className="text-blue-600" />
            <h3 className="font-semibold text-slate-800 dark:text-slate-100">Standardized SPFx Project Structure</h3>
          </div>
          <CodeBlock code={SPFX_TREE} language="directory tree" copyKey="spfx-tree" copiedKey={copiedKey} onCopy={onCopy} />
        </Card>
      )}

      {tab === "flows" && (
        <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
          {POWER_AUTOMATE_FLOWS.map((flow) => (
            <Card key={flow.id} className="p-5">
              <div className="mb-2 flex items-center gap-2">
                <Workflow size={16} className="text-indigo-600" />
                <h3 className="font-semibold text-slate-800 dark:text-slate-100">{flow.name}</h3>
              </div>
              <p className="mb-3 text-sm text-slate-500 dark:text-slate-400">{flow.description}</p>
              <div className="mb-3 rounded-lg bg-slate-50 p-2.5 text-xs dark:bg-slate-800/60">
                <span className="font-bold uppercase tracking-wide text-slate-400">Trigger: </span>
                <span className="text-slate-600 dark:text-slate-300">{flow.trigger}</span>
              </div>
              <div>
                <p className="mb-1.5 text-xs font-bold uppercase tracking-wide text-slate-400">Actions</p>
                <ol className="ml-4 list-decimal space-y-1 text-xs text-slate-600 dark:text-slate-300">
                  {flow.actions.map((a, i) => (
                    <li key={i}>{a}</li>
                  ))}
                </ol>
              </div>
            </Card>
          ))}
        </div>
      )}

      {tab === "connect" && (
        <div className="space-y-4">
          <Card className="p-5">
            <div className="mb-3 flex items-center gap-2">
              <Cloud size={16} className="text-blue-600" />
              <h3 className="font-semibold text-slate-800 dark:text-slate-100">Connection Status</h3>
            </div>
            <div className="flex flex-wrap items-center gap-3">
              <SyncBadge status={spStatus} error={spError} />
              <p className="text-sm text-slate-500 dark:text-slate-400">
                {spStatus === "connected" && "Work Log & Issue Tracker are reading and writing live SharePoint list data."}
                {spStatus === "connecting" && "Loading Work Log & Issue Tracker items from SharePoint…"}
                {spStatus === "error" && (spError || "Something went wrong talking to SharePoint — showing local data instead.")}
                {(spStatus === "idle" || spStatus === "unconfigured") &&
                  "Sign in with a Microsoft account to connect — until then, Work Log & Issue Tracker use local demo data only."}
              </p>
            </div>
          </Card>

          <Card className="p-5">
            <h3 className="mb-3 font-semibold text-slate-800 dark:text-slate-100">1. Create the SharePoint site and lists</h3>
            <p className="text-sm text-slate-500 dark:text-slate-400">
              In your SharePoint Online tenant, create (or pick) a site, then add <strong>BenchProjects</strong>,{" "}
              <strong>BenchWorkLogs</strong>, and <strong>BenchIssueTracker</strong> custom lists with the columns shown in the{" "}
              <em>SharePoint List Schemas</em> tab; create BenchProjects first, since BenchWorkLogs' ProjectName Lookup column
              points at it. The other four lists are optional for now — only Projects, Work Log, and Issue Tracker read/write
              live data in this build.
            </p>
            <p className="mt-3 text-sm text-slate-500 dark:text-slate-400">
              For BenchProjects' <strong>ProjectManager</strong>, <strong>TechnicalLead</strong>, and <strong>TeamMembers</strong>{" "}
              columns, and for BenchIssueTracker's <strong>AssignedTo</strong> column, only people who've opened the site at
              least once show up as pickable — those dropdowns are populated from the site's actual member list, not free text.
            </p>
            <p className="mt-3 text-sm text-slate-500 dark:text-slate-400">
              Add a <strong>BenchAdministration</strong> list with <strong>UserName</strong>, <strong>UserEmail</strong>, and{" "}
              <strong>Access_Level</strong> (Admin / Manager / User) columns. You don't need to pre-populate it — the first
              time anyone signs in with Microsoft, a row is created for them automatically with Access_Level set to{" "}
              <strong>User</strong>; a manual "Sign In" (see step 5) instead just checks whether a row with that email already
              exists, so an admin adds those rows by hand. Only <strong>Admin</strong> sees this Architecture (Admin) section,
              so after your own row appears, open the list directly in SharePoint and change your Access_Level to Admin — this
              fails closed, so until you do that (or if this list doesn't exist yet), nobody sees this section, yourself
              included.
            </p>
          </Card>

          <Card className="p-5">
            <h3 className="mb-3 font-semibold text-slate-800 dark:text-slate-100">2. Register an Entra ID app</h3>
            <ol className="ml-4 list-decimal space-y-1.5 text-sm text-slate-600 dark:text-slate-300">
              <li>
                In the <a className="text-blue-600 underline dark:text-blue-400" href="https://entra.microsoft.com" target="_blank" rel="noopener noreferrer">Microsoft Entra admin center</a>, create a new App Registration.
              </li>
              <li>
                Under Authentication, add a <strong>Single-page application</strong> platform with redirect URI{" "}
                <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">{typeof window !== "undefined" ? window.location.origin : "http://localhost:5173"}</code>.
              </li>
              <li>
                Under API permissions, add Microsoft Graph delegated permission{" "}
                <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">Sites.ReadWrite.All</code> — this covers
                everything except the Issue Tracker's Assignee field.
              </li>
              <li>
                Also under API permissions, click <strong>Add a permission → APIs my organization uses</strong>, search for{" "}
                <strong>SharePoint</strong>, choose <strong>Delegated permissions</strong>, and add{" "}
                <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">AllSites.Read</code>. This is what
                resolves Person columns — the Issue Tracker's Assignee and BenchProjects' ProjectManager / TechnicalLead /
                TeamMembers — to a real site user; it's a separate API from Graph, so it needs its own entry here.
              </li>
              <li>Click <strong>Grant admin consent for [your org]</strong> to approve both permissions at once (needs admin rights).</li>
              <li>Copy the Application (client) ID and Directory (tenant) ID from the Overview page.</li>
            </ol>
          </Card>

          <Card className="p-5">
            <h3 className="mb-3 font-semibold text-slate-800 dark:text-slate-100">3. Configure Bench</h3>
            <p className="mb-3 text-sm text-slate-500 dark:text-slate-400">
              Copy <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">.env.example</code> to{" "}
              <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">.env</code>, fill in these values, then restart the dev server:
            </p>
            <div className="overflow-hidden rounded-xl border border-slate-200 dark:border-slate-800">
              <table className="w-full text-left text-sm">
                <thead className="bg-slate-50 text-xs uppercase tracking-wide text-slate-500 dark:bg-slate-900/60 dark:text-slate-400">
                  <tr>
                    <th className="px-3 py-2 font-medium">Variable</th>
                    <th className="px-3 py-2 font-medium">Current Value</th>
                    <th className="px-3 py-2 font-medium">Notes</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                  {envRows.map((row) => (
                    <tr key={row.key}>
                      <td className="px-3 py-2 font-mono text-xs font-semibold text-blue-700 dark:text-blue-400">{row.key}</td>
                      <td className="px-3 py-2 font-mono text-xs text-slate-600 dark:text-slate-300">
                        {row.value ? row.value : <span className="text-slate-400">not set</span>}
                      </td>
                      <td className="px-3 py-2 text-xs text-slate-500 dark:text-slate-400">{row.note}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>

          <Card className="p-5">
            <h3 className="mb-3 font-semibold text-slate-800 dark:text-slate-100">4. Sign in</h3>
            <p className="text-sm text-slate-500 dark:text-slate-400">
              Sign out and back in using the <strong>Microsoft</strong> button on the login screen. Once all four variables above are
              set, that button opens a real Microsoft account picker instead of the demo flow, and Work Log &amp; Issue Tracker switch
              to live SharePoint data automatically.
            </p>
          </Card>

          <Card className="p-5">
            <h3 className="mb-3 font-semibold text-slate-800 dark:text-slate-100">5. (Optional) Manual Sign In against SharePoint</h3>
            <p className="text-sm text-slate-500 dark:text-slate-400">
              The email "Sign In" form has no Microsoft credential of its own, so checking BenchAdministration needs an
              app-only (client credentials) Graph token — and Microsoft's identity platform flatly refuses to issue that kind
              of token to any request carrying a browser Origin header, specifically to keep app secrets out of client-side
              code. So this can't run in the browser at all; a tiny backend in the{" "}
              <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">/server</code> folder of this
              project does the token exchange and the list lookup instead, and the browser just calls it. "Create Account" is
              paused for now — Sign In (allow if the email has a row, no password check) is the only manual auth path in use.
            </p>
            <ol className="mt-3 ml-4 list-decimal space-y-1.5 text-sm text-slate-600 dark:text-slate-300">
              <li>
                Register an Entra app (confidential client) and add an application permission of{" "}
                <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">Sites.ReadWrite.All</code> on
                Microsoft Graph. Also click <strong>Add a permission → APIs my organization uses</strong>, search for{" "}
                <strong>Office 365 SharePoint Online</strong>, choose <strong>Application permissions</strong>, and add{" "}
                <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">Sites.FullControl.All</code> — this
                is what lets this server resolve BenchProjects' ProjectManager / TechnicalLead / TeamMembers to real site
                users; skip it and those People fields on the Add Project form just fall back to the local demo names. Click{" "}
                <strong>Grant admin consent</strong> to approve both permissions. Under Certificates &amp; secrets, create a
                client secret and copy its value immediately — it's shown once.
              </li>
              <li>
                In <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">server/</code>, copy{" "}
                <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">.env.example</code> to{" "}
                <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">.env</code>, fill in{" "}
                <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">GRAPH_TENANT_ID</code>,{" "}
                <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">GRAPH_CLIENT_ID</code>,{" "}
                <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">GRAPH_CLIENT_SECRET</code>,{" "}
                <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">SP_HOSTNAME</code>, and{" "}
                <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">SP_SITE_PATH</code>, then run{" "}
                <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">npm install &amp;&amp; npm start</code>{" "}
                (defaults to port 8787).
              </li>
              <li>
                Set <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">VITE_AUTH_API_URL</code> here
                to that server's base URL (e.g.{" "}
                <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">http://localhost:8787</code>{" "}
                for local dev), then restart the frontend dev server.
              </li>
            </ol>
            <p className="mt-3 text-sm text-slate-500 dark:text-slate-400">
              Once set, Sign In sends the entered email to that server, which checks it against BenchAdministration and lets
              the person in if a row exists — there's no password to enter. Leave{" "}
              <code className="rounded bg-slate-100 px-1 py-0.5 text-xs dark:bg-slate-800">VITE_AUTH_API_URL</code> unset and
              Sign In shows a clear "not connected" error instead of silently doing nothing.
            </p>
          </Card>
        </div>
      )}
    </div>
  );
}

/* ----------------------------------------------------------------------- */
/* Root application                                                         */
/* ----------------------------------------------------------------------- */

export default function Bench() {
  const [theme, setTheme] = useState(() => {
    try {
      return localStorage.getItem("bench-theme") || "light";
    } catch {
      return "light";
    }
  });
  const [user, setUser] = useState(() => {
    try {
      const raw = localStorage.getItem("bench-user");
      return raw ? JSON.parse(raw) : null;
    } catch {
      return null;
    }
  });
  const [authMode, setAuthMode] = useState("signin");
  const [ssoLoading, setSsoLoading] = useState(null);
  const [ssoError, setSsoError] = useState("");
  const [formLoading, setFormLoading] = useState(false);
  const [authError, setAuthError] = useState("");

  const spConfigured = isSharePointConfigured();
  const [spStatus, setSpStatus] = useState(spConfigured ? "idle" : "unconfigured");
  const [spError, setSpError] = useState("");

  const [activeView, setActiveView] = useState("dashboard");
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [globalSearch, setGlobalSearch] = useState("");

  const [workLogs, setWorkLogs] = useState(SEED_WORK_LOGS);
  const [issues, setIssues] = useState(SEED_ISSUES);
  const [projects, setProjects] = useState(SEED_PROJECTS);
  const [projectOptions, setProjectOptions] = useState(PROJECTS);
  const [assigneeOptions, setAssigneeOptions] = useState(ASSIGNEES);
  const [accessLevel, setAccessLevel] = useState(null);

  const [clippyOpen, setClippyOpen] = useState(false);
  const [clippyMood, setClippyMood] = useState("idle");
  const [showHint, setShowHint] = useState(true);
  const [messages, setMessages] = useState([
    { id: "m0", role: "clippy", text: "Hi, I'm Clippy, your Bench AI Copilot! Ask me to summarize your work logs, explain a Graph/SPFx/Power Platform pattern, or help draft resolution notes." },
  ]);

  const [copiedKey, onCopy] = useCopyToClipboard();

  useEffect(() => {
    try {
      localStorage.setItem("bench-theme", theme);
    } catch {
      /* ignore */
    }
  }, [theme]);

  useEffect(() => {
    const t = setTimeout(() => setShowHint(false), 6000);
    return () => clearTimeout(t);
  }, []);

  // Runs once on every app load. On a normal load this resolves null and does
  // nothing; if this load is the browser landing back from signInWithMicrosoft's
  // redirect, it resolves the signed-in account and completes the sign-in.
  useEffect(() => {
    if (!isMsalConfigured()) return;
    let cancelled = false;
    (async () => {
      try {
        const account = await completeMicrosoftSignIn();
        if (cancelled || !account) return;
        setUser({
          name: account.name || account.username,
          email: account.username,
          role: "Microsoft 365 User",
          provider: "microsoft",
        });
      } catch (err) {
        if (!cancelled) setSsoError(err?.message || "Microsoft sign-in failed.");
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    try {
      if (user) localStorage.setItem("bench-user", JSON.stringify(user));
      else localStorage.removeItem("bench-user");
    } catch {
      /* ignore */
    }
  }, [user]);

  const handleSsoLogin = async (provider) => {
    setSsoError("");
    if (provider === "microsoft" && isMsalConfigured()) {
      setSsoLoading("microsoft");
      try {
        // Navigates the whole page away to Microsoft's sign-in; on success this
        // call never returns here at all — the result is picked up by the
        // startup effect below after the browser lands back on this app.
        await signInWithMicrosoft();
      } catch (err) {
        setSsoError(err?.message || "Microsoft sign-in failed. Please try again.");
        setSsoLoading(null);
      }
      return;
    }
    // Google (and Microsoft when no App Registration is configured) use a
    // simulated sign-in — see the Architecture (Admin) setup notes for wiring
    // up real Microsoft Entra ID sign-in against SharePoint.
    setSsoLoading(provider);
    setTimeout(() => {
      setUser(AUTH_PROFILES[provider]);
      setSsoLoading(null);
    }, 1000);
  };

  const handleEmailLogin = async ({ name, email }) => {
    setFormLoading(true);
    setAuthError("");
    try {
      const { name: storedName } = await verifyManualLogin({ email });
      setUser({ name: storedName || name, email, role: "Team Member", provider: "email" });
    } catch (err) {
      setAuthError(err?.message || "Sign-in failed. Please try again.");
    } finally {
      setFormLoading(false);
    }
  };

  // Manual "Create Account" is paused for now (see LoginView) — Sign In alone
  // checks the entered email against BenchAdministration. registerManualUser
  // in lib/manualAuth.js still does the write and is ready to wire back in
  // behind a signup form later.

  const handleGuestLogin = () => {
    setUser({ name: "Guest User", email: "guest@bench.local", role: "Guest", provider: "guest" });
  };

  const handleLogout = () => {
    if (user?.provider === "microsoft" && isMsalConfigured()) {
      // logoutRedirect navigates the whole page away, so nothing after this
      // call is guaranteed to run — clear the persisted session first so the
      // app doesn't just log the same user back in from localStorage once
      // Microsoft redirects back here.
      try {
        localStorage.removeItem("bench-user");
      } catch {
        /* ignore */
      }
      signOutMicrosoft();
      return;
    }
    setUser(null);
    setActiveView("dashboard");
    setClippyOpen(false);
    setSpStatus(spConfigured ? "idle" : "unconfigured");
    setSpError("");
    setProjects(SEED_PROJECTS);
    setProjectOptions(PROJECTS);
    setAssigneeOptions(ASSIGNEES);
    setAccessLevel(null);
  };

  // Microsoft sign-in syncs Work Log, Issue Tracker, Projects and Assignees
  // via a delegated Graph token. Manual (email) sign-in has no such token, so
  // it only syncs Work Log + Projects, through this server's app-only
  // credential (see manualWorkLogService.js) — Issue Tracker's Assignee
  // resolution needs a delegated SharePoint REST call that path doesn't have.
  const isMicrosoftSync = user?.provider === "microsoft" && spConfigured;
  const isManualSync = user?.provider === "email" && isManualWorkLogAvailable();

  useEffect(() => {
    if (!isMicrosoftSync && !isManualSync) return;
    let cancelled = false;
    (async () => {
      setSpStatus("connecting");
      setSpError("");
      try {
        if (isMicrosoftSync) {
          const [logs, tickets, projectItems, assignees] = await Promise.all([
            fetchWorkLogs(),
            fetchIssues(),
            fetchProjects(),
            fetchAssignees(),
          ]);
          if (cancelled) return;
          setWorkLogs(logs);
          setIssues(tickets);
          if (projectItems.length) {
            setProjects(projectItems);
            setProjectOptions(projectItems.map((p) => p.title));
          }
          if (assignees.length) setAssigneeOptions(assignees);
        } else {
          const [logs, projectItems] = await Promise.all([fetchManualWorkLogs(), fetchManualProjects()]);
          if (cancelled) return;
          setWorkLogs(logs);
          if (projectItems.length) {
            setProjects(projectItems);
            setProjectOptions(projectItems.map((p) => p.title));
          }
          // Separate from the load above: the People fields on the Add
          // Project form are a nice-to-have, and the extra SharePoint
          // app-only permission they need (see Architecture (Admin) →
          // Connect SharePoint) might not be granted yet — that shouldn't
          // block Work Log / Projects from loading.
          try {
            const siteUsers = await fetchManualSiteUsers();
            if (!cancelled && siteUsers.length) setAssigneeOptions(siteUsers);
          } catch {
            /* People fields fall back to the local ASSIGNEES list */
          }
        }
        setSpStatus("connected");
      } catch (err) {
        if (cancelled) return;
        setSpError(err?.message || "Could not load data from SharePoint.");
        setSpStatus("error");
      }
      if (!isMicrosoftSync) return;
      // Independent of the sync above: BenchAdministration only controls
      // whether the Architecture (Admin) nav item is shown, so a problem here
      // (list not created yet, user not listed) shouldn't block Work Log /
      // Issue Tracker sync — it just defaults to the non-admin "User" level.
      try {
        // Prefer the real mailbox address (Graph /me's "mail") over the sign-in
        // UPN we got from MSAL — they're often different, and BenchAdministration
        // is more likely to be populated with the former.
        let lookupEmail = user.email;
        let lookupName = user.name;
        try {
          const profile = await fetchMyProfile();
          if (profile.email) lookupEmail = profile.email;
          if (profile.name) lookupName = profile.name;
        } catch {
          /* fall back to the account info already in user */
        }
        const level = await ensureAdministrationEntry({ name: lookupName, email: lookupEmail });
        if (!cancelled) setAccessLevel(level);
      } catch {
        if (!cancelled) setAccessLevel("User");
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [isMicrosoftSync, isManualSync]);

  const addWorkLog = async (entry) => {
    const withSubmitter = { ...entry, submittedBy: user?.email || "" };
    if (spStatus !== "connected") {
      setWorkLogs((prev) => [withSubmitter, ...prev]);
      return;
    }
    setWorkLogs((prev) => [withSubmitter, ...prev]);
    try {
      const saved = await (isMicrosoftSync ? createWorkLog(withSubmitter) : createManualWorkLog(withSubmitter));
      setWorkLogs((prev) => [saved, ...prev.filter((l) => l.id !== entry.id)]);
    } catch (err) {
      setSpError(err?.message || "Failed to save the entry to SharePoint.");
      setSpStatus("error");
    }
  };

  const deleteWorkLog = async (id) => {
    setWorkLogs((prev) => prev.filter((l) => l.id !== id));
    if (spStatus !== "connected") return;
    try {
      await (isMicrosoftSync ? removeWorkLog(id) : removeManualWorkLog(id));
    } catch (err) {
      setSpError(err?.message || "Failed to delete the entry from SharePoint.");
      setSpStatus("error");
    }
  };

  const addProject = async (project) => {
    const { title } = project;
    if (spStatus !== "connected") {
      setProjects((prev) => [...prev, { ...project, id: `local-${Date.now()}` }]);
      setProjectOptions((prev) => (prev.includes(title) ? prev : [...prev, title]));
      return;
    }
    const tempId = `pending-${Date.now()}`;
    setProjects((prev) => [...prev, { ...project, id: tempId }]);
    setProjectOptions((prev) => (prev.includes(title) ? prev : [...prev, title]));
    try {
      const saved = await (isMicrosoftSync ? createProject(project) : createManualProject(project));
      setProjects((prev) => prev.map((p) => (p.id === tempId ? saved : p)));
    } catch (err) {
      setProjects((prev) => prev.filter((p) => p.id !== tempId));
      setProjectOptions((prev) => prev.filter((t) => t !== title));
      setSpError(err?.message || "Failed to save the project to SharePoint.");
      setSpStatus("error");
    }
  };

  const editProject = async (id, project) => {
    const previous = projects.find((p) => p.id === id);
    setProjects((prev) => prev.map((p) => (p.id === id ? { ...project, id } : p)));
    setProjectOptions((prev) => {
      const withoutOldTitle = previous && previous.title !== project.title ? prev.filter((t) => t !== previous.title) : prev;
      return withoutOldTitle.includes(project.title) ? withoutOldTitle : [...withoutOldTitle, project.title];
    });
    if (spStatus !== "connected") return;
    try {
      await (isMicrosoftSync ? updateProject(id, project) : updateManualProject(id, project));
    } catch (err) {
      if (previous) setProjects((prev) => prev.map((p) => (p.id === id ? previous : p)));
      setSpError(err?.message || "Failed to update the project in SharePoint.");
      setSpStatus("error");
    }
  };

  const deleteProject = async (id) => {
    const removed = projects.find((p) => p.id === id);
    setProjects((prev) => prev.filter((p) => p.id !== id));
    if (removed) setProjectOptions((prev) => prev.filter((title) => title !== removed.title));
    if (spStatus !== "connected") return;
    try {
      await (isMicrosoftSync ? removeProject(id) : removeManualProject(id));
    } catch (err) {
      setSpError(err?.message || "Failed to delete the project from SharePoint.");
      setSpStatus("error");
    }
  };

  const addIssue = async (issue) => {
    if (spStatus !== "connected") {
      setIssues((prev) => [...prev, { ...issue, id: `is-${Date.now()}` }]);
      return;
    }
    const tempId = `is-${Date.now()}`;
    setIssues((prev) => [...prev, { ...issue, id: tempId }]);
    try {
      const saved = await createIssue(issue);
      setIssues((prev) => [...prev.filter((t) => t.id !== tempId), saved]);
    } catch (err) {
      setSpError(err?.message || "Failed to save the ticket to SharePoint.");
      setSpStatus("error");
    }
  };

  const moveIssueColumn = async (id, column) => {
    setIssues((prev) => prev.map((t) => (t.id === id ? { ...t, column } : t)));
    if (spStatus !== "connected") return;
    try {
      await moveIssue(id, column);
    } catch (err) {
      setSpError(err?.message || "Failed to update the ticket in SharePoint.");
      setSpStatus("error");
    }
  };

  const sendClippyMessage = (text) => {
    setShowHint(false);
    setMessages((prev) => [...prev, { id: `u-${Date.now()}`, role: "user", text }]);
    setClippyMood("thinking");
    setTimeout(() => {
      const reply = generateClippyReply(text, { workLogs, challenges: CHALLENGES });
      setMessages((prev) => [...prev, { id: `c-${Date.now()}`, role: "clippy", text: reply }]);
      setClippyMood("celebrate");
      setTimeout(() => setClippyMood("idle"), 1400);
    }, 900);
  };

  const toggleTheme = () => setTheme((t) => (t === "dark" ? "light" : "dark"));

  // Local/demo sessions (email, Google, guest) aren't tied to a real identity,
  // so they keep full access for exploration. A real Microsoft sign-in is
  // gated by BenchAdministration's Access_Level — only "Admin" sees the
  // Architecture (Admin) section; "Manager" and "User" don't.
  const isAdmin = user?.provider !== "microsoft" || accessLevel === "Admin";
  const visibleNavItems = isAdmin ? NAV_ITEMS : NAV_ITEMS.filter((item) => item.key !== "admin");

  useEffect(() => {
    if (activeView === "admin" && !isAdmin) setActiveView("dashboard");
  }, [activeView, isAdmin]);

  return (
    <div className={theme === "dark" ? "dark" : ""}>
      {!user ? (
        <LoginView
          mode={authMode}
          onSsoLogin={handleSsoLogin}
          ssoLoading={ssoLoading}
          ssoError={ssoError}
          onEmailLogin={handleEmailLogin}
          formLoading={formLoading}
          authError={authError}
          onGuestLogin={handleGuestLogin}
          theme={theme}
          onToggleTheme={toggleTheme}
        />
      ) : (
        <div className="min-h-screen bg-slate-100 text-slate-900 transition-colors duration-300 dark:bg-slate-950 dark:text-slate-100">
          <Header
            theme={theme}
            onToggleTheme={toggleTheme}
            onToggleMobileMenu={() => setMobileMenuOpen(true)}
            search={globalSearch}
            onSearch={setGlobalSearch}
            user={user}
            onLogout={handleLogout}
            accessLevel={accessLevel}
          />

          <div className="flex w-full">
            <Sidebar
              activeView={activeView}
              onNavigate={setActiveView}
              collapsed={sidebarCollapsed}
              onToggleCollapse={() => setSidebarCollapsed((c) => !c)}
              items={visibleNavItems}
            />
            <MobileNav
              open={mobileMenuOpen}
              onClose={() => setMobileMenuOpen(false)}
              activeView={activeView}
              onNavigate={setActiveView}
              items={visibleNavItems}
            />

            <main className="min-w-0 flex-1 p-4 sm:p-6 lg:p-8">
              {activeView === "dashboard" && (
                <DashboardView workLogs={workLogs} challenges={CHALLENGES} learningItems={LEARNING_ITEMS} onNavigate={setActiveView} user={user} />
              )}
              {activeView === "worklog" && (
                <WorkLogView
                  workLogs={workLogs}
                  onAdd={addWorkLog}
                  onDelete={deleteWorkLog}
                  syncBadge={<SyncBadge status={spStatus} error={spError} />}
                  projectOptions={projectOptions}
                />
              )}
              {activeView === "projects" && (
                <ProjectsView
                  projects={projects}
                  onAdd={addProject}
                  onEdit={editProject}
                  onDelete={deleteProject}
                  syncBadge={<SyncBadge status={spStatus} error={spError} />}
                  peopleOptions={assigneeOptions}
                />
              )}
              {activeView === "knowledge" && <KnowledgeBaseView copiedKey={copiedKey} onCopy={onCopy} />}
              {activeView === "challenges" && <ChallengesView />}
              {activeView === "learning" && <LearningCenterView />}
              {activeView === "components" && <ComponentsLibraryView copiedKey={copiedKey} onCopy={onCopy} />}
              {activeView === "issues" && (
                <IssueTrackerView
                  issues={issues}
                  onCreateIssue={addIssue}
                  onMoveIssue={moveIssueColumn}
                  syncBadge={<SyncBadge status={spStatus} error={spError} />}
                  assigneeOptions={assigneeOptions}
                />
              )}
              {activeView === "admin" && isAdmin && (
                <ArchitectureAdminView copiedKey={copiedKey} onCopy={onCopy} spStatus={spStatus} spError={spError} />
              )}
            </main>
          </div>

          <ClippyWidget open={clippyOpen} onToggle={() => setClippyOpen((o) => !o)} mood={clippyMood} showHint={showHint} />
          <ChatDrawer open={clippyOpen} onClose={() => setClippyOpen(false)} messages={messages} onSend={sendClippyMessage} mood={clippyMood} />
        </div>
      )}
    </div>
  );
}
