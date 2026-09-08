import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'

export async function connectClient(server: McpServer): Promise<Client> {
  const [clientTransport, serverTransport] =
    InMemoryTransport.createLinkedPair()
  const client = new Client({ name: 'test', version: '0.0.0' })
  await Promise.all([
    client.connect(clientTransport),
    server.connect(serverTransport),
  ])
  return client
}

export function textOf(result: unknown): string {
  if (!result || typeof result !== 'object') return ''
  const content = (result as { content?: unknown }).content
  if (!Array.isArray(content)) return ''
  return content
    .filter(
      (item): item is { type: string; text?: string } =>
        !!item &&
        typeof item === 'object' &&
        (item as { type?: unknown }).type === 'text',
    )
    .map(item => item.text ?? '')
    .join('\n')
}

export function isToolError(result: unknown): boolean {
  return (
    !!result &&
    typeof result === 'object' &&
    (result as { isError?: unknown }).isError === true
  )
}
