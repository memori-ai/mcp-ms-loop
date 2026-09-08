# MCP client configuration

Canonical stdio configurations for `@memori.ai/mcp-ms-loop`.

All values below are placeholders. Store credentials in the private encrypted
configuration of the MCP host, never in a versioned file.

## AIsuru MCP Gateway

- `server_type`: `generic`
- command: `npx -y @memori.ai/mcp-ms-loop`
- runtime: Node >= 24

Parameters:

```json
{
  "TENANT_ID": "<MICROSOFT_ENTRA_TENANT_ID>",
  "CLIENT_ID": "<MICROSOFT_ENTRA_CLIENT_ID>",
  "CLIENT_SECRET": "<MICROSOFT_ENTRA_CLIENT_SECRET>",
  "LOOP_MAX_CONTENT_CHARS": "100000"
}
```

Suggested `prompt_lista`:

> Use this MCP to discover accessible Microsoft Loop workspaces and files, then
> read Loop content by sharing URL or driveId/itemId. Inspect the available
> tools before calling them.

Suggested `prompt_esegui`:

> Prefer the default text output. Request HTML or both formats only when
> structure is necessary. Page workspace and file lists with offset and limit,
> and request only the content size needed. Never request or reveal Microsoft
> credentials.

After configuration, **Fetch Tools** must show:

- `getLoopByShareUrl`
- `getLoopByItemId`
- `listLoopContainers`
- `listLoopFilesInDrive`

## Claude Desktop

```json
{
  "mcpServers": {
    "ms-loop": {
      "command": "npx",
      "args": ["-y", "@memori.ai/mcp-ms-loop"],
      "env": {
        "TENANT_ID": "<MICROSOFT_ENTRA_TENANT_ID>",
        "CLIENT_ID": "<MICROSOFT_ENTRA_CLIENT_ID>",
        "CLIENT_SECRET": "<MICROSOFT_ENTRA_CLIENT_SECRET>"
      }
    }
  }
}
```

Use the same `mcpServers` entry in Cursor or LM Studio.

For a local checkout, run `bun run build`, then replace command and args with:

```json
{
  "command": "node",
  "args": ["/ABSOLUTE/PATH/TO/mcp-ms-loop/build/cli.js"]
}
```

## Claude Code

Configure the three required variables through the environment or the private
secret mechanism available to the host, then add:

```sh
claude mcp add ms-loop -- npx -y @memori.ai/mcp-ms-loop
```

Do not place `CLIENT_SECRET` directly in a shell command because shell history
may retain it.
