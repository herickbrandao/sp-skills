# SharePoint REST MCP Server

Standalone MCP server for SharePoint list and library management. All SharePoint data calls use SharePoint REST endpoints under `/_api`; Microsoft Graph is not used.

## Tools exposed to the AI

- `sp_get_sites` — list configured sites (tenant-wide discovery is not generally available through ordinary site-scoped REST permissions).
- `sp_get_access` — validate the current Microsoft identity and access to a site.
- `sp_get_lists` — list site lists and document libraries.
- `sp_get_columns` — inspect list field names, types, and settings.
- `sp_get_items` — read list items with OData filters and continuation paging.
- `sp_create_item` — create a list item.
- `sp_update_item` — update a list item using ETags.
- `sp_delete_item` — delete or recycle a list item.
- `sp_batch` — submit bounded, validated SharePoint REST batch requests.
- `sp_rest` — make a constrained SharePoint REST call for supported endpoints not covered above.

## Authentication

The AI never receives or supplies an access token. The server obtains a SharePoint-audience bearer token using the configured Microsoft credential provider, adds it to REST requests, and obtains/caches a per-site request digest for writes.

Supported modes:

- `azure_cli`: use `AzureCliCredential`; sign in before starting the server with `az login`.
- `device_code`: use `DeviceCodeCredential`; the first server request prints a Microsoft device-code sign-in prompt to stderr. Register a public client app and grant the required SharePoint delegated permissions.
- `client_secret`: use `ClientSecretCredential` for service-to-service automation. This is app-only access, not the signed-in user's delegated access; assign least-privilege SharePoint application permissions and admin consent. Keep the secret in the process environment/secret store, never in source control.

For every mode, the token scope is `https://{SP_TENANT_HOST}/.default` (SharePoint audience, not Graph). Configure `SP_TENANT_HOST`, then copy `.env.example` to `.env` or set the variables in the process environment. The server does not load `.env` automatically, so secrets should be injected by the MCP host or a secret manager.

## Run

Requires Node.js 20 or newer.

```sh
npm install
```

After installation, use the standalone skill instructions in `skills/sharepoint-rest/SKILL.md` with the agent. If your agent host requires skills inside a central skills directory, copy that folder there.

Set authentication/site variables, then configure the MCP host to launch:

```json
{
  "mcpServers": {
    "sharepoint-rest": {
      "command": "node",
      "args": ["/absolute/path/to/sharepoint-rest-mcp/src/server.js"],
      "env": {
        "SP_AUTH_MODE": "azure_cli",
        "SP_TENANT_HOST": "contoso.sharepoint.com",
        "SP_SITE_URLS": "https://contoso.sharepoint.com/sites/Operations"
      }
    }
  }
}
```

The server uses stdio transport. Logs and device-code prompts go to stderr; stdout is reserved for MCP messages.

## Agent instructions

Use `skills/sharepoint-rest/SKILL.md` as the agent-facing skill entry point. The skill directs the agent to discover sites, validate access, inspect list schemas before writes, use dedicated tools for common operations, follow continuation tokens, and never request credentials from the user in chat.

## Notes

- `sp_get_sites` returns the configured catalog. Tenant-wide site discovery requires separately provisioned discovery/administrative access and is intentionally not implied by site-level access.
- For `sp_update_item` and `sp_delete_item`, the server fetches the current ETag if omitted. `if_match_any` is an explicit opt-in to `If-Match: *`.
- `sp_rest` and `sp_batch` accept only same-site relative `_api/` paths; authentication headers are server-managed.
- Microsoft identity and SharePoint permissions still govern all operations. The MCP server does not grant permissions.

## References

- [SharePoint REST: lists and list items](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest)
- [SharePoint REST OData queries](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/use-odata-query-operations-in-sharepoint-rest-requests)
- [SharePoint REST batch requests](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/make-batch-requests-with-the-rest-apis)
- [Microsoft identity platform and OAuth](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)
