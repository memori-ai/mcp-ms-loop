import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js'

const transport = new StdioClientTransport({
  command: process.execPath,
  args: ['build/cli.js'],
  env: {
    ...process.env,
    TENANT_ID: 'smoke-tenant',
    CLIENT_ID: 'smoke-client',
    CLIENT_SECRET: 'smoke-secret',
  },
  stderr: 'pipe',
})
const client = new Client({
  name: 'mcp-ms-loop-node-smoke',
  version: '1.0.0',
})

try {
  await client.connect(transport)
  const response = await client.listTools()
  const names = response.tools.map(tool => tool.name).sort()
  const expected = [
    'getLoopByItemId',
    'getLoopByShareUrl',
    'listLoopContainers',
    'listLoopFilesInDrive',
  ]
  if (JSON.stringify(names) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected MCP tools: ${names.join(', ')}`)
  }
  process.stdout.write('Node MCP smoke OK: 4 tools\n')
} finally {
  await client.close()
}
