import type { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { createLoopMcpServer, type LoopServiceApi } from '../src/server.js'
import { LoopError } from '../src/services/errors.js'

import { connectClient, isToolError, textOf } from './mcp-test-utils.js'

const TOOL_NAMES = [
  'getLoopByItemId',
  'getLoopByShareUrl',
  'listLoopContainers',
  'listLoopFilesInDrive',
]

function serviceStub(
  calls: Array<{ name: string; args: unknown[] }>,
): LoopServiceApi {
  return {
    async getLoopByShareUrl(...args) {
      calls.push({ name: 'getLoopByShareUrl', args })
      return {
        name: 'Example.loop',
        driveId: 'drive',
        itemId: 'item',
        webUrl: null,
        lastModified: null,
        size: 10,
        mimeType: 'application/octet-stream',
        text: 'Example',
        textTruncated: false,
      }
    },
    async getLoopByItemId(...args) {
      calls.push({ name: 'getLoopByItemId', args })
      return {
        name: 'Example.loop',
        driveId: 'drive',
        itemId: 'item',
        webUrl: null,
        lastModified: null,
        size: 10,
        mimeType: 'application/octet-stream',
        text: 'Example',
        textTruncated: false,
      }
    },
    async listLoopContainers(...args) {
      calls.push({ name: 'listLoopContainers', args })
      return {
        items: [],
        pagina: {
          offset: 0,
          limite: 20,
          restituiti: 0,
          haAltri: false,
          totaleScoperti: 0,
          risultatiRicercaParziali: false,
        },
      }
    },
    async listLoopFilesInDrive(...args) {
      calls.push({ name: 'listLoopFilesInDrive', args })
      return {
        items: [],
        pagina: {
          offset: 0,
          limite: 50,
          restituiti: 0,
          haAltri: false,
        },
      }
    },
  }
}

describe('Loop MCP contract', () => {
  let client: Client
  let calls: Array<{ name: string; args: unknown[] }>

  beforeEach(async () => {
    calls = []
    client = await connectClient(
      createLoopMcpServer({ service: serviceStub(calls) }),
    )
  })

  afterEach(async () => {
    await client.close()
  })

  it('keeps exactly the four stable tool names', async () => {
    const { tools } = await client.listTools()
    expect(tools.map(tool => tool.name).sort()).toEqual(TOOL_NAMES)
  })

  it('publishes read-only open-world annotations', async () => {
    const { tools } = await client.listTools()
    for (const tool of tools) {
      expect(tool.description?.trim().length).toBeGreaterThan(10)
      expect(tool.annotations?.readOnlyHint).toBe(true)
      expect(tool.annotations?.openWorldHint).toBe(true)
    }
  })

  it('keeps tools/list below five kilobytes without secrets', async () => {
    const response = await client.listTools()
    const serialized = JSON.stringify(response.tools)
    expect(Buffer.byteLength(serialized)).toBeLessThanOrEqual(5_000)
    expect(serialized).not.toContain('client-secret-value')
  })

  it('dispatches all four happy paths with bounded defaults', async () => {
    const results = await Promise.all([
      client.callTool({
        name: 'getLoopByShareUrl',
        arguments: {
          shareUrl: 'https://example.invalid/shared',
          contentMode: 'text',
          maxChars: 1_000,
        },
      }),
      client.callTool({
        name: 'getLoopByItemId',
        arguments: {
          driveId: 'drive',
          itemId: 'item',
          includeText: true,
        },
      }),
      client.callTool({
        name: 'listLoopContainers',
        arguments: {},
      }),
      client.callTool({
        name: 'listLoopFilesInDrive',
        arguments: { driveId: 'drive', path: 'LoopAppData' },
      }),
    ])

    expect(results.every(result => !isToolError(result))).toBe(true)
    expect(calls.map(call => call.name).sort()).toEqual(TOOL_NAMES)
    expect(calls.find(call => call.name === 'getLoopByShareUrl')?.args).toEqual(
      [
        'https://example.invalid/shared',
        { contentMode: 'text', maxChars: 1_000 },
      ],
    )
    expect(
      calls.find(call => call.name === 'listLoopContainers')?.args,
    ).toEqual([{ offset: 0, limit: 20 }])
    expect(
      calls.find(call => call.name === 'listLoopFilesInDrive')?.args,
    ).toEqual(['drive', 'LoopAppData', { offset: 0, limit: 50 }])
  })

  it('returns schema errors for invalid input', async () => {
    const result = await client.callTool({
      name: 'getLoopByShareUrl',
      arguments: { shareUrl: 'not-a-url' },
    })
    expect(isToolError(result)).toBe(true)
  })
})

describe('Loop MCP expected errors', () => {
  it('returns sanitized structured tool errors', async () => {
    const service = serviceStub([])
    service.getLoopByItemId = async () => {
      throw new LoopError(
        'Microsoft Graph denied the request.',
        'GRAPH_REQUEST_FAILED',
        403,
      )
    }
    const client = await connectClient(createLoopMcpServer({ service }))
    try {
      const result = await client.callTool({
        name: 'getLoopByItemId',
        arguments: { driveId: 'drive', itemId: 'item' },
      })
      expect(isToolError(result)).toBe(true)
      expect(JSON.parse(textOf(result))).toEqual({
        error: true,
        code: 'GRAPH_REQUEST_FAILED',
        message: 'Microsoft Graph denied the request.',
        status: 403,
      })
    } finally {
      await client.close()
    }
  })
})
