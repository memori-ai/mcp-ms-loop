# AGENTS.md — Microsoft Loop MCP

Instructions for engineers and coding agents working on
`@memori.ai/mcp-ms-loop`.

## Goal and non-negotiables

This package is a read-only Microsoft Loop MCP server over stdio.

- Development uses Bun >= 1.4.
- Production uses Node >= 24 through
  `npx -y @memori.ai/mcp-ms-loop`.
- The published entrypoint is `build/cli.js`; source must remain
  Node-compatible.
- `TENANT_ID`, `CLIENT_ID` and `CLIENT_SECRET` come only from environment
  variables.
- Never print credentials, bearer tokens, Axios objects, tenant ids, item ids,
  URLs, Loop content, or raw Microsoft Graph errors.
- Stdout is exclusively the MCP transport. Safe operational logs use stderr.
- Expected failures return compact structured JSON with `isError: true`.
- All responses must respect the content and pagination budgets.

## Stable public contract

Do not rename or remove these tools:

- `getLoopByShareUrl`
- `getLoopByItemId`
- `listLoopContainers`
- `listLoopFilesInDrive`

All four tools are read-only and open-world. Keep
`readOnlyHint: true` and `openWorldHint: true`.

`contentMode` controls content output:

- `text` is the compact default.
- `html` returns HTML only.
- `both` shares one total budget between HTML and text.
- The deprecated `includeText` alias remains supported:
  `false` means `html`, `true` means `both`; `contentMode` wins.

List tools return `{items,pagina}` and must retain bounded `offset`/`limit`.

## Environment contract

- `TENANT_ID` — required Microsoft Entra tenant id.
- `CLIENT_ID` — required application client id.
- `CLIENT_SECRET` — required application secret.
- `LOOP_MAX_CONTENT_CHARS` — optional total response-content budget; default
  100000, maximum 400000.

Any change must be reflected in `src/lib/env.ts`, README and
`docs/MCP_CLIENTS.md`.

## Layout

- `src/cli.ts` — validates env and connects stdio.
- `src/server.ts` — registers the four tools.
- `src/services/` — OAuth and Microsoft Graph domain logic.
- `src/lib/` — typed env, response and limit helpers.
- `test/` — credential-free in-memory and mocked Graph tests.
- `build/` — generated and published output.

## Graph rules

- Preserve standard share encoding and SharePoint Embedded `nav` parsing.
- Keep the Search API region fallback.
- Select only required Graph fields.
- Follow `@odata.nextLink` only until the requested page plus one item is
  known.
- Validate drive-relative paths and encode path segments.
- Never return unbounded arrays or duplicate full HTML and text by default.
- Do not add Graph write permissions or mutative tools without an explicit
  design and administrative write gate.

## Required checks

```bash
bun install --frozen-lockfile
bun run lint:typeaware
bun run typecheck
bun run test
bun run format:check
bun run build
node --check build/cli.js
bun audit --production
npm pack --dry-run
```

Tests must not require a real tenant, credential, or network request.
Publishing with `npm publish` is supported on Node >= 24: `prepack` invokes
`npm run build`, so the publish step does not require a globally installed Bun
binary.

## Gateway checklist

- `server_type: generic`
- command: `npx -y @memori.ai/mcp-ms-loop`
- parameters match the environment contract exactly
- Node runtime is at least version 24
- Fetch Tools returns the four real tools
- credentials are stored as encrypted secrets

Do not add `.env` files, certificates, Graph output, tenant fixtures, logs,
generated build artifacts, or package tarballs to source control.
