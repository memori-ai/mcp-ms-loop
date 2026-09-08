#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'

import { loadEnv } from './lib/env.js'
import { createLoopMcpServer, VERSION } from './server.js'

async function main(): Promise<void> {
  let env: ReturnType<typeof loadEnv>
  try {
    env = loadEnv()
  } catch (error) {
    process.stderr.write(
      `[mcp-ms-loop] configuration error: ${
        error instanceof Error ? error.message : String(error)
      }\n`,
    )
    process.exit(1)
  }

  const server = createLoopMcpServer({
    credentials: {
      tenantId: env.TENANT_ID,
      clientId: env.CLIENT_ID,
      clientSecret: env.CLIENT_SECRET,
    },
    maxContentChars: env.LOOP_MAX_CONTENT_CHARS,
  })
  await server.connect(new StdioServerTransport())
  process.stderr.write(`[mcp-ms-loop] v${VERSION} connected over stdio\n`)

  const shutdown = async () => {
    await server.close()
    process.exit(0)
  }
  process.once('SIGINT', () => void shutdown())
  process.once('SIGTERM', () => void shutdown())
}

main().catch(() => {
  process.stderr.write('[mcp-ms-loop] fatal startup error\n')
  process.exit(1)
})
