import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { SharePointClient, boundedInt, listPath, parseSiteCatalog, validateODataPath } from "./sharepoint.js";

const sp = new SharePointClient();
const server = new McpServer({ name: "sharepoint-rest-mcp", version: "1.0.0" });
const siteUrlSchema = z.string().url().describe("Absolute HTTPS URL of a site in the configured SharePoint tenant.");
const listSelector = {
  list_id: z.string().uuid().optional().describe("List GUID. Preferred over list_title."),
  list_title: z.string().min(1).optional().describe("Display title of the list or library."),
};
const jsonObject = z.record(z.unknown());

function result(data) {
  return { content: [{ type: "text", text: JSON.stringify(data) }], structuredContent: data };
}

function fail(error) {
  const details = error?.details ?? { message: error instanceof Error ? error.message : String(error) };
  return { content: [{ type: "text", text: JSON.stringify({ error: details }) }], structuredContent: { error: details }, isError: true };
}

function register(name, config, handler) {
  server.registerTool(name, config, async (input) => {
    try { return result(await handler(input)); }
    catch (error) { return fail(error); }
  });
}

function requireList(selector) {
  if (!!selector.list_id === !!selector.list_title) throw new Error("Provide exactly one of list_id or list_title.");
  return listPath(selector);
}

function normalizeField(field) {
  return {
    id: field.Id ?? field.ID ?? null,
    internal_name: field.InternalName ?? field.StaticName ?? null,
    display_name: field.Title ?? null,
    type: field.TypeAsString ?? field.FieldTypeKind ?? null,
    required: field.Required ?? false,
    read_only: field.ReadOnlyField ?? false,
    hidden: field.Hidden ?? false,
    choices: field.Choices?.results ?? field.Choices ?? null,
    default_value: field.DefaultValue ?? null,
    lookup_list: field.LookupList ?? null,
    lookup_field: field.LookupField ?? null,
    allow_multiple_values: field.AllowMultipleValues ?? null,
  };
}

register("sp_get_sites", {
  title: "Get SharePoint sites",
  description: "List sites in the server's configured SharePoint site catalog. This does not use Microsoft Graph or imply tenant-wide discovery permissions.",
  inputSchema: {
    query: z.string().optional().describe("Optional case-insensitive filter over configured site URL or title."),
    limit: z.number().int().min(1).max(200).default(50),
    continuation_token: z.string().optional().describe("Offset token returned by the previous page."),
  },
  annotations: { readOnlyHint: true, openWorldHint: false },
}, async ({ query, limit, continuation_token }) => {
  let sites = parseSiteCatalog().map((site) => ({ ...site, title: site.title || site.url }));
  if (query) sites = sites.filter((site) => `${site.title} ${site.url}`.toLowerCase().includes(query.toLowerCase()));
  const offset = continuation_token ? Number(continuation_token) : 0;
  if (!Number.isInteger(offset) || offset < 0) throw new Error("Invalid continuation_token.");
  const page = sites.slice(offset, offset + limit);
  return { sites: page, next_continuation_token: offset + page.length < sites.length ? String(offset + page.length) : null };
});

register("sp_get_access", {
  title: "Check SharePoint access",
  description: "Verify the configured Microsoft identity and read access to a SharePoint site. Never returns access tokens or secrets.",
  inputSchema: { site_url: siteUrlSchema },
  annotations: { readOnlyHint: true, openWorldHint: true },
}, async ({ site_url }) => {
  try {
    const response = await sp.request(site_url, "_api/web", { query: { "$select": "Id,Title,Url,CurrentUser/Title,CurrentUser/Email", "$expand": "CurrentUser" } });
    const web = response.data ?? {};
    return {
      authenticated: true,
      has_access: true,
      site_url: sp.siteUrl(site_url),
      site_title: web.Title ?? null,
      site_id: web.Id ?? null,
      user: web.CurrentUser ? { display_name: web.CurrentUser.Title ?? null, email: web.CurrentUser.Email ?? null } : null,
      message: "The current Microsoft identity can access this SharePoint site.",
    };
  } catch (error) {
    const status = error?.details?.status;
    return {
      authenticated: status !== 401,
      has_access: false,
      site_url: sp.siteUrl(site_url),
      message: status === 401 ? "Microsoft authentication is missing or expired. Sign in through the configured credential provider." :
        status === 403 ? "The current Microsoft identity is signed in but does not have permission to this site." : error.message,
      error: error?.details ?? null,
    };
  }
});

register("sp_get_lists", {
  title: "Get SharePoint lists",
  description: "List visible SharePoint lists and document libraries on a site, with paging.",
  inputSchema: {
    site_url: siteUrlSchema,
    include_hidden: z.boolean().default(false),
    limit: z.number().int().min(1).max(500).default(100),
    continuation_token: z.string().optional(),
  },
  annotations: { readOnlyHint: true, openWorldHint: true },
}, async ({ site_url, include_hidden, limit, continuation_token }) => {
  const page = await sp.paged(site_url, "_api/web/lists", {
    "$select": "Id,Title,Description,BaseTemplate,BaseType,Hidden,ItemCount,RootFolder/ServerRelativeUrl,ListItemEntityTypeFullName",
    "$expand": "RootFolder",
    "$filter": include_hidden ? undefined : "Hidden eq false",
  }, limit, continuation_token);
  return {
    lists: page.items.map((list) => ({
      id: list.Id,
      title: list.Title,
      description: list.Description ?? "",
      base_template: list.BaseTemplate ?? null,
      base_type: list.BaseType ?? null,
      kind: list.BaseType === 1 ? "document_library" : "list",
      hidden: list.Hidden ?? false,
      item_count: list.ItemCount ?? null,
      root_folder_url: list.RootFolder?.ServerRelativeUrl ?? null,
      list_item_entity_type: list.ListItemEntityTypeFullName ?? null,
    })),
    next_continuation_token: page.next_continuation_token,
    count: page.count,
  };
});

register("sp_get_columns", {
  title: "Get SharePoint columns",
  description: "Read a list's column internal names, types, validation flags, choices, defaults, and lookup settings.",
  inputSchema: {
    site_url: siteUrlSchema,
    ...listSelector,
    include_hidden: z.boolean().default(false),
    limit: z.number().int().min(1).max(500).default(100),
    continuation_token: z.string().optional(),
  },
  annotations: { readOnlyHint: true, openWorldHint: true },
}, async ({ site_url, list_id, list_title, include_hidden, limit, continuation_token }) => {
  const base = `_api/web/${requireList({ list_id, list_title })}/fields`;
  const page = await sp.paged(site_url, base, {
    "$select": "Id,InternalName,StaticName,Title,TypeAsString,FieldTypeKind,Required,ReadOnlyField,Hidden,Choices,DefaultValue,LookupList,LookupField,AllowMultipleValues",
    "$filter": include_hidden ? undefined : "Hidden eq false",
  }, limit, continuation_token);
  return { columns: page.items.map(normalizeField), next_continuation_token: page.next_continuation_token, count: page.count };
});

register("sp_get_items", {
  title: "Get SharePoint list items",
  description: "Query list items through SharePoint REST with OData select, expand, filter, orderby, and continuation paging.",
  inputSchema: {
    site_url: siteUrlSchema,
    ...listSelector,
    item_id: z.number().int().positive().optional(),
    select: z.array(z.string()).max(100).optional(),
    expand: z.array(z.string()).max(50).optional(),
    filter: z.string().optional().describe("SharePoint OData $filter expression."),
    orderby: z.string().optional().describe("SharePoint OData $orderby expression."),
    page_size: z.number().int().min(1).max(500).default(100),
    continuation_token: z.string().optional(),
  },
  annotations: { readOnlyHint: true, openWorldHint: true },
}, async ({ site_url, list_id, list_title, item_id, select, expand, filter, orderby, page_size, continuation_token }) => {
  const base = `_api/web/${requireList({ list_id, list_title })}/items`;
  if (item_id) {
    boundedInt(item_id, undefined, 1, Number.MAX_SAFE_INTEGER);
    const response = await sp.request(site_url, `${base}(${item_id})`, { query: {
      "$select": select?.length ? select.join(",") : undefined,
      "$expand": expand?.length ? expand.join(",") : undefined,
    } });
    return { items: response.data ? [response.data] : [], next_continuation_token: null, count: response.data ? 1 : 0 };
  }
  const page = await sp.paged(site_url, base, {
    "$select": select?.length ? select.join(",") : undefined,
    "$expand": expand?.length ? expand.join(",") : undefined,
    "$filter": filter,
    "$orderby": orderby,
  }, page_size, continuation_token);
  return page;
});

register("sp_create_item", {
  title: "Create SharePoint list item",
  description: "Create one item in a SharePoint list. Field keys must be SharePoint internal names.",
  inputSchema: { site_url: siteUrlSchema, ...listSelector, fields: jsonObject.describe("Writable SharePoint field internal names and values.") },
  annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
}, async ({ site_url, list_id, list_title, fields }) => {
  const list = requireList({ list_id, list_title });
  const metadata = await sp.request(site_url, `_api/web/${list}`, { query: { "$select": "ListItemEntityTypeFullName" } });
  const type = metadata.data?.ListItemEntityTypeFullName ?? metadata.raw?.d?.ListItemEntityTypeFullName;
  const body = type ? { __metadata: { type }, ...fields } : fields;
  const response = await sp.write(site_url, `_api/web/${list}/items`, { method: "POST", body });
  const item = response.data ?? {};
  return { item_id: item.Id ?? item.ID ?? item.id ?? null, item, status: response.status };
});

register("sp_update_item", {
  title: "Update SharePoint list item",
  description: "Update fields on one item. Uses the current ETag unless supplied; if_match_any explicitly permits overwriting concurrent changes.",
  inputSchema: {
    site_url: siteUrlSchema,
    ...listSelector,
    item_id: z.number().int().positive(),
    fields: jsonObject,
    etag: z.string().optional(),
    if_match_any: z.boolean().default(false),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true },
}, async ({ site_url, list_id, list_title, item_id, fields, etag, if_match_any }) => {
  const list = requireList({ list_id, list_title });
  const base = `_api/web/${list}/items`;
  const match = if_match_any ? "*" : (etag || await sp.ensureEtag(site_url, base, item_id));
  const response = await sp.write(site_url, `${base}(${boundedInt(item_id, undefined, 1, Number.MAX_SAFE_INTEGER)})`, {
    method: "POST",
    headers: { "X-HTTP-Method": "MERGE", "If-Match": match },
    body: fields,
  });
  return { item_id, updated: response.status >= 200 && response.status < 300, etag: response.etag, status: response.status };
});

register("sp_delete_item", {
  title: "Delete or recycle SharePoint list item",
  description: "Delete one exact list item, or move it to the recycle bin when requested. Requires explicit site, list, and item ID.",
  inputSchema: {
    site_url: siteUrlSchema,
    ...listSelector,
    item_id: z.number().int().positive(),
    etag: z.string().optional(),
    if_match_any: z.boolean().default(false),
    recycle: z.boolean().default(false),
  },
  annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true },
}, async ({ site_url, list_id, list_title, item_id, etag, if_match_any, recycle }) => {
  const list = requireList({ list_id, list_title });
  const base = `_api/web/${list}/items`;
  const id = boundedInt(item_id, undefined, 1, Number.MAX_SAFE_INTEGER);
  const match = if_match_any ? "*" : (etag || await sp.ensureEtag(site_url, base, id));
  const response = recycle
    ? await sp.write(site_url, `${base}(${id})/recycle()`, { method: "POST", body: "" })
    : await sp.write(site_url, `${base}(${id})`, { method: "POST", headers: { "X-HTTP-Method": "DELETE", "If-Match": match } });
  return { item_id: id, deleted: !recycle, recycled: recycle, status: response.status };
});

const batchMax = boundedInt(process.env.SP_MAX_BATCH_REQUESTS || "20", 20, 1, 50);
register("sp_batch", {
  title: "Run SharePoint REST batch",
  description: "Submit a bounded set of same-site REST requests to _api/$batch. Batch calls may partially succeed; inspect each subresponse.",
  inputSchema: {
    site_url: siteUrlSchema,
    requests: z.array(z.object({
      action: z.enum(["GET", "POST", "UPDATE", "PATCH", "DELETE"]),
      path: z.string().startsWith("_api/"),
      item: z.unknown().optional(),
      if_match: z.string().optional(),
    })).min(1).max(batchMax),
  },
  annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true },
}, async ({ site_url, requests }) => {
  const site = sp.siteUrl(site_url);
  const digest = await sp.digest(site);
  const boundary = `batch_${crypto.randomUUID()}`;
  const changeBoundary = `changeset_${crypto.randomUUID()}`;
  const reads = [];
  const writes = [];
  for (const request of requests) {
    const path = validateODataPath(request.path);
    if (path.includes("?") && /[\r\n]/.test(path)) throw new Error("Invalid batch path.");
    const action = request.action === "UPDATE" ? "PATCH" : request.action;
    (action === "GET" ? reads : writes).push({ ...request, path, action });
  }
  const lines = [];
  if (writes.length) {
    lines.push(`--${boundary}`, `Content-Type: multipart/mixed; boundary="${changeBoundary}"`, "");
    for (const request of writes) {
      lines.push(`--${changeBoundary}`, "Content-Type: application/http", "Content-Transfer-Encoding: binary", "", `${request.action} ${site}/${request.path.replace(/^_api\//, "_api/")} HTTP/1.1`);
      if (request.if_match || request.action === "DELETE" || request.action === "PATCH") lines.push(`If-Match: ${request.if_match ?? "*"}`);
      if (request.item !== undefined) lines.push("Content-Type: application/json;odata=nometadata", "", JSON.stringify(request.item));
      else lines.push("");
    }
    lines.push(`--${changeBoundary}--`);
  }
  for (const request of reads) lines.push(`--${boundary}`, "Content-Type: application/http", "Content-Transfer-Encoding: binary", "", `GET ${site}/${request.path} HTTP/1.1`, "Accept: application/json;odata=nometadata", "");
  lines.push(`--${boundary}--`, "");
  const response = await sp.request(site, "_api/$batch", { method: "POST", body: lines.join("\r\n"), headers: {
    "Content-Type": `multipart/mixed; boundary="${boundary}"`,
    "X-RequestDigest": digest,
  } });
  return { status: response.status, raw_response: response.data ?? response.raw, request_count: requests.length, note: "Inspect each multipart response; individual operations may fail independently." };
});

register("sp_rest", {
  title: "Call SharePoint REST endpoint",
  description: "Call a constrained SharePoint REST endpoint not covered by a dedicated tool. Only same-site relative _api paths are accepted; authentication and request digest are server-managed.",
  inputSchema: {
    site_url: siteUrlSchema,
    method: z.enum(["GET", "POST", "PUT", "PATCH", "DELETE"]),
    path: z.string().startsWith("_api/"),
    query: z.record(z.union([z.string(), z.number(), z.boolean(), z.array(z.string())])).optional(),
    headers: z.record(z.string()).optional().describe("Optional allowlisted headers. Authentication/cookie/host headers are forbidden."),
    body: z.union([jsonObject, z.string()]).optional(),
  },
  annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true },
}, async ({ site_url, method, path, query, headers = {}, body }) => {
  const forbidden = new Set(["authorization", "cookie", "set-cookie", "host", "origin", "content-length", "x-requestdigest"]);
  for (const key of Object.keys(headers)) if (forbidden.has(key.toLowerCase())) throw new Error(`Header '${key}' is managed or forbidden.`);
  const allowedHeaders = new Set(["accept", "content-type", "if-match", "x-http-method", "prefer"]);
  for (const key of Object.keys(headers)) if (!allowedHeaders.has(key.toLowerCase())) throw new Error(`Header '${key}' is not allowed.`);
  const safePath = validateODataPath(path);
  const options = { method, query, headers, body };
  const response = method === "GET" ? await sp.request(site_url, safePath, options) : await sp.write(site_url, safePath, options);
  return { status: response.status, etag: response.etag, data: response.data };
});

const transport = new StdioServerTransport();
await server.connect(transport);
