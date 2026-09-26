const { app } = require("@azure/functions");
const {
  isGraphAppConfigured,
  getAppToken,
  graphFetch,
  resolveSiteId,
  resolveListId,
  getColumnNameMap,
} = require("../graphAppClient");

const LIST_NAME = "BenchAdministration";
const ACCESS_LEVELS = ["Admin", "Manager", "User"];
const DEFAULT_ACCESS_LEVEL = "User";
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || "*";

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
}

app.http("registerUser", {
  methods: ["POST", "OPTIONS"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    if (request.method === "OPTIONS") {
      return { status: 204, headers: corsHeaders() };
    }

    if (!isGraphAppConfigured()) {
      return {
        status: 501,
        headers: corsHeaders(),
        jsonBody: { error: "This function isn't configured yet — set GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, SP_HOSTNAME, SP_SITE_PATH." },
      };
    }

    let body;
    try {
      body = await request.json();
    } catch {
      return { status: 400, headers: corsHeaders(), jsonBody: { error: "Invalid JSON body." } };
    }

    const name = String(body?.name || "").trim();
    const email = String(body?.email || "").trim();
    if (!email) {
      return { status: 400, headers: corsHeaders(), jsonBody: { error: "email is required." } };
    }

    try {
      const token = await getAppToken();
      const siteId = await resolveSiteId(token);
      const listId = await resolveListId(token, siteId, LIST_NAME);
      const columnMap = await getColumnNameMap(token, siteId, listId);
      const nameField = columnMap.get("UserName") || "UserName";
      const emailField = columnMap.get("UserEmail") || "UserEmail";
      const levelField = columnMap.get("Access_Level") || "Access_Level";

      const existingItems = await graphFetch(token, `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500`);
      const normalized = email.toLowerCase();
      const existing = (existingItems.value || []).find(
        (item) => String(item.fields?.[emailField] || "").trim().toLowerCase() === normalized
      );

      let accessLevel = DEFAULT_ACCESS_LEVEL;
      if (existing) {
        const level = existing.fields?.[levelField];
        accessLevel = ACCESS_LEVELS.includes(level) ? level : DEFAULT_ACCESS_LEVEL;
      } else {
        await graphFetch(token, `/sites/${siteId}/lists/${listId}/items`, {
          method: "POST",
          body: JSON.stringify({
            fields: {
              [nameField]: name || email,
              [emailField]: email,
              [levelField]: DEFAULT_ACCESS_LEVEL,
            },
          }),
        });
      }

      return { status: 200, headers: corsHeaders(), jsonBody: { accessLevel } };
    } catch (err) {
      context.error(err);
      return { status: 500, headers: corsHeaders(), jsonBody: { error: err.message || "Failed to register user." } };
    }
  },
});
