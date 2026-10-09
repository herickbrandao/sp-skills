import {
  AzureCliCredential,
  ClientSecretCredential,
  DeviceCodeCredential,
} from "@azure/identity";

const MAX_PAGE_SIZE = 500;
const MAX_BODY_BYTES = 5 * 1024 * 1024;
const ALLOWED_METHODS = new Set(["GET", "POST", "PUT", "PATCH", "DELETE"]);

function env(name, fallback = "") {
  return (process.env[name] ?? fallback).trim();
}

function parseSiteCatalog() {
  return env("SP_SITE_URLS")
    .split(/[;\n]/)
    .map((value) => value.trim())
    .filter(Boolean)
    .map((url) => ({ url, title: "" }));
}

function parseResponseBody(text, contentType) {
  if (!text) return null;
  if (contentType.includes("json") || /^[\s]*[\[{]/.test(text)) {
    try { return JSON.parse(text); } catch { /* Keep non-JSON error payload below. */ }
  }
  return text;
}

function sharePointError(payload, status) {
  const error = payload?.error ?? payload?.["odata.error"];
  const message = error?.message?.value ?? error?.message ?? payload?.message ??
    (typeof payload === "string" ? payload.slice(0, 1000) : `SharePoint REST request failed with HTTP ${status}`);
  const code = error?.code ?? payload?.code ?? null;
  return { status, code, message: String(message), correlation_id: null };
}

function unwrap(payload) {
  if (payload && typeof payload === "object" && "d" in payload) {
    const data = payload.d;
    if (data && typeof data === "object" && Array.isArray(data.results)) return data.results;
    return data;
  }
  return payload;
}

function odataString(value) {
  return String(value).replaceAll("'", "''");
}

function listPath({ list_id, list_title }) {
  if (list_id) {
    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(list_id)) {
      throw new Error("list_id must be a valid GUID.");
    }
    return `lists(guid'${list_id}')`;
  }
  if (!list_title?.trim()) throw new Error("Provide list_id or list_title.");
  return `lists/GetByTitle('${encodeURIComponent(odataString(list_title.trim()))}')`;
}

function boundedInt(value, fallback, min = 1, max = MAX_PAGE_SIZE) {
  const num = Number(value ?? fallback);
  if (!Number.isInteger(num) || num < min || num > max) throw new Error(`Value must be an integer from ${min} to ${max}.`);
  return num;
}

function validateODataPath(path) {
  if (typeof path !== "string" || !path.startsWith("_api/")) throw new Error("path must be a relative SharePoint path beginning with _api/.");
  if (path.startsWith("//") || path.includes("\\") || /[\r\n]/.test(path)) throw new Error("Invalid SharePoint REST path.");
  const decoded = decodeURIComponent(path);
  if (decoded.split("/").some((part) => part === "..") || /^[a-z][a-z\d+.-]*:/i.test(decoded)) throw new Error("Path traversal and absolute URLs are not allowed.");
  return path.replace(/^_api\//, "_api/");
}

function validateQueryValue(value) {
  if (typeof value === "string" && /[\r\n]/.test(value)) throw new Error("Query values cannot contain line breaks.");
  return value;
}

export class SharePointClient {
  constructor() {
    this.tenantHost = env("SP_TENANT_HOST").toLowerCase();
    this.tenantId = env("SP_TENANT_ID");
    this.clientId = env("SP_CLIENT_ID");
    this.clientSecret = env("SP_CLIENT_SECRET");
    this.authMode = env("SP_AUTH_MODE", "azure_cli").toLowerCase();
    this.sites = parseSiteCatalog();
    this.digestCache = new Map();
    this.credential = null;
    this.devicePrompted = false;
  }

  validateConfiguration() {
    if (!this.tenantHost || !/^(?=.{1,253}$)[a-z0-9](?:[a-z0-9.-]*[a-z0-9])?$/i.test(this.tenantHost)) {
      throw new Error("SP_TENANT_HOST must be the SharePoint Online hostname, such as contoso.sharepoint.com.");
    }
    if (!["azure_cli", "device_code", "client_secret"].includes(this.authMode)) throw new Error("SP_AUTH_MODE must be azure_cli, device_code, or client_secret.");
    if (this.authMode !== "azure_cli" && (!this.tenantId || !this.clientId)) throw new Error("SP_TENANT_ID and SP_CLIENT_ID are required for this auth mode.");
    if (this.authMode === "client_secret" && !this.clientSecret) throw new Error("SP_CLIENT_SECRET is required for client_secret mode.");
  }

  getCredential() {
    this.validateConfiguration();
    if (this.credential) return this.credential;
    if (this.authMode === "azure_cli") this.credential = new AzureCliCredential();
    else if (this.authMode === "device_code") {
      this.credential = new DeviceCodeCredential({
        tenantId: this.tenantId,
        clientId: this.clientId,
        userPromptCallback: (info) => {
          this.devicePrompted = true;
          process.stderr.write(`${info.message}\n`);
        },
      });
    } else this.credential = new ClientSecretCredential(this.tenantId, this.clientId, this.clientSecret);
    return this.credential;
  }

  async accessToken() {
    const credential = this.getCredential();
    const token = await credential.getToken(`https://${this.tenantHost}/.default`);
    if (!token?.token) throw new Error("Microsoft identity provider returned no SharePoint access token.");
    return token.token;
  }

  siteUrl(input) {
    if (typeof input !== "string" || !input.trim()) throw new Error("site_url is required.");
    let url;
    try { url = new URL(input); } catch { throw new Error("site_url must be an absolute HTTPS URL."); }
    if (url.protocol !== "https:" || url.hostname.toLowerCase() !== this.tenantHost) throw new Error(`site_url must use the configured SharePoint host https://${this.tenantHost}.`);
    if (url.username || url.password || url.search || url.hash) throw new Error("site_url cannot include credentials, query parameters, or fragments.");
    return url.origin + url.pathname.replace(/\/$/, "");
  }

  async request(siteInput, path, options = {}) {
    const site = this.siteUrl(siteInput);
    const method = (options.method ?? "GET").toUpperCase();
    if (!ALLOWED_METHODS.has(method)) throw new Error("Unsupported HTTP method.");
    const rel = validateODataPath(path);
    const url = new URL(`${site}/${rel}`);
    if (options.query) {
      for (const [key, rawValue] of Object.entries(options.query)) {
        if (!/^\$?[A-Za-z][A-Za-z\d]*$/.test(key)) throw new Error(`Invalid OData query key: ${key}`);
        const value = validateQueryValue(rawValue);
        if (Array.isArray(value)) url.searchParams.set(key, value.join(","));
        else if (value !== undefined && value !== null) url.searchParams.set(key, String(value));
      }
    }
    const headers = {
      Accept: "application/json;odata=nometadata",
      ...(options.body !== undefined ? { "Content-Type": "application/json;odata=nometadata" } : {}),
      ...(options.headers ?? {}),
      Authorization: `Bearer ${await this.accessToken()}`,
    };
    let body = options.body;
    if (body !== undefined && typeof body !== "string" && !(body instanceof Uint8Array)) body = JSON.stringify(body);
    if (typeof body === "string" && Buffer.byteLength(body) > MAX_BODY_BYTES) throw new Error("Request body exceeds the configured size limit.");
    const response = await fetch(url, {
      method,
      headers,
      body,
      signal: AbortSignal.timeout(30000),
      redirect: "error",
    });
    const text = await response.text();
    if (Buffer.byteLength(text) > MAX_BODY_BYTES) throw new Error("SharePoint response exceeds the 5 MB safety limit. Narrow the query or page size.");
    const payload = parseResponseBody(text, response.headers.get("content-type") ?? "");
    if (!response.ok) {
      const err = sharePointError(payload, response.status);
      err.correlation_id = response.headers.get("sprequestguid") ?? response.headers.get("request-id");
      if (response.status === 429 || response.status === 503) err.retry_after = response.headers.get("retry-after");
      const error = new Error(err.message);
      error.details = err;
      throw error;
    }
    return { status: response.status, etag: response.headers.get("etag"), data: unwrap(payload), raw: payload };
  }

  async digest(siteInput) {
    const site = this.siteUrl(siteInput);
    const cached = this.digestCache.get(site);
    if (cached && cached.expiresAt > Date.now() + 30000) return cached.value;
    const result = await this.request(site, "_api/contextinfo", { method: "POST", body: "" });
    const info = result.data?.GetContextWebInformation ?? result.raw?.d?.GetContextWebInformation;
    const value = info?.FormDigestValue;
    if (!value) throw new Error("SharePoint did not return a form digest from _api/contextinfo.");
    const timeout = Number(info.FormDigestTimeoutSeconds) || 1500;
    this.digestCache.set(site, { value, expiresAt: Date.now() + Math.max(timeout - 30, 30) * 1000 });
    return value;
  }

  async write(site, path, options = {}) {
    const digest = await this.digest(site);
    return this.request(site, path, {
      ...options,
      headers: { ...(options.headers ?? {}), "X-RequestDigest": digest },
    });
  }

  async paged(site, path, query, pageSize, continuationToken) {
    const pageLimit = boundedInt(pageSize, 100, 1, MAX_PAGE_SIZE);
    let result;
    if (continuationToken) {
      let next;
      try { next = new URL(continuationToken); } catch { throw new Error("Invalid continuation_token."); }
      const expectedSite = new URL(this.siteUrl(site));
      if (next.protocol !== "https:" || next.hostname.toLowerCase() !== this.tenantHost || next.origin !== expectedSite.origin || !next.pathname.startsWith(`${expectedSite.pathname}/_api/`)) {
        throw new Error("Continuation token must point to the selected site's SharePoint REST endpoint.");
      }
      result = await this.request(site, next.pathname.slice(expectedSite.pathname.length + 1) + next.search);
    } else {
      result = await this.request(site, path, { query: { ...query, "$top": pageLimit } });
    }
    const data = result.data;
    const items = Array.isArray(data) ? data : (data?.value ?? data?.results ?? []);
    const nextLink = data?.__next ?? data?.["@odata.nextLink"] ?? result.raw?.d?.__next ?? null;
    return { items, next_continuation_token: nextLink, count: items.length };
  }

  async ensureEtag(site, basePath, itemId) {
    const result = await this.request(site, `${basePath}(${boundedInt(itemId, undefined, 1, Number.MAX_SAFE_INTEGER)})`, { query: { "$select": "Id" } });
    return result.etag ?? result.raw?.__metadata?.etag ?? result.raw?.d?.__metadata?.etag ?? "*";
  }
}

export { boundedInt, listPath, odataString, validateODataPath, parseSiteCatalog };
