import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'

import { DEFAULT_LIMITS } from './lib/limits.js'
import { toolError, toolJson } from './lib/tools.js'
import type { LoopCredentials } from './services/authService.js'
import {
  LoopService,
  type ContentMode,
  type ContentOptions,
  type PageRequest,
} from './services/loopService.js'

export const VERSION = '1.0.0'

export type LoopServiceApi = Pick<
  LoopService,
  | 'getLoopByShareUrl'
  | 'getLoopByItemId'
  | 'listLoopContainers'
  | 'listLoopFilesInDrive'
>

export type LoopMcpConfig = {
  credentials?: LoopCredentials
  maxContentChars?: number
  service?: LoopServiceApi
}

export function createLoopMcpServer(config: LoopMcpConfig): McpServer {
  const maxContentChars =
    config.maxContentChars ?? DEFAULT_LIMITS.maxContentChars
  const service =
    config.service ??
    new LoopService(
      config.credentials ?? {
        tenantId: '',
        clientId: '',
        clientSecret: '',
      },
      maxContentChars,
    )
  const server = new McpServer(
    { name: 'mcp-ms-loop', version: VERSION },
    { capabilities: { tools: {} } },
  )

  const annotations = {
    readOnlyHint: true,
    openWorldHint: true,
  } as const
  const contentFields = {
    contentMode: z
      .enum(['text', 'html', 'both'])
      .optional()
      .describe('Output: text (default), html, or both.'),
    includeText: z
      .boolean()
      .optional()
      .describe('Deprecated alias: true=both, false=html. contentMode wins.'),
    maxChars: z
      .number()
      .int()
      .min(DEFAULT_LIMITS.minimumContentChars)
      .max(maxContentChars)
      .optional()
      .describe(
        `Serialized response budget in UTF-8 bytes; max ${maxContentChars}.`,
      ),
  }
  const pageFields = (defaultLimit: number) => ({
    offset: z
      .number()
      .int()
      .min(0)
      .default(0)
      .describe('Zero-based result offset.'),
    limit: z
      .number()
      .int()
      .min(1)
      .max(DEFAULT_LIMITS.maxPageSize)
      .default(defaultLimit)
      .describe(
        `Page size; default ${defaultLimit}, max ${DEFAULT_LIMITS.maxPageSize}.`,
      ),
  })
  const contentOptions = (args: {
    contentMode?: ContentMode
    includeText?: boolean
    maxChars?: number
  }): ContentOptions => ({
    ...(args.contentMode === undefined
      ? {}
      : { contentMode: args.contentMode }),
    ...(args.includeText === undefined
      ? {}
      : { includeText: args.includeText }),
    ...(args.maxChars === undefined ? {} : { maxChars: args.maxChars }),
  })
  const pageRequest = (args: {
    offset: number
    limit: number
  }): PageRequest => ({
    offset: args.offset,
    limit: args.limit,
  })

  server.registerTool(
    'getLoopByShareUrl',
    {
      description:
        'Read a Microsoft Loop file from a OneDrive, SharePoint, or Loop app sharing URL.',
      inputSchema: {
        shareUrl: z
          .string()
          .url()
          .max(4_096)
          .describe('Full HTTP(S) sharing URL of the Loop file.'),
        ...contentFields,
      },
      annotations,
    },
    async args => {
      try {
        return toolJson(
          await service.getLoopByShareUrl(args.shareUrl, contentOptions(args)),
        )
      } catch (error) {
        return toolError(error)
      }
    },
  )

  server.registerTool(
    'getLoopByItemId',
    {
      description:
        'Read a Microsoft Loop file directly from its driveId and itemId.',
      inputSchema: {
        driveId: z
          .string()
          .min(1)
          .max(512)
          .describe('Drive containing the Loop file.'),
        itemId: z.string().min(1).max(512).describe('Loop file item id.'),
        ...contentFields,
      },
      annotations,
    },
    async args => {
      try {
        return toolJson(
          await service.getLoopByItemId(
            args.driveId,
            args.itemId,
            contentOptions(args),
          ),
        )
      } catch (error) {
        return toolError(error)
      }
    },
  )

  server.registerTool(
    'listLoopContainers',
    {
      description:
        'Discover accessible Microsoft Loop workspaces and their driveIds.',
      inputSchema: pageFields(DEFAULT_LIMITS.defaultContainerPageSize),
      annotations,
    },
    async args => {
      try {
        return toolJson(await service.listLoopContainers(pageRequest(args)))
      } catch (error) {
        return toolError(error)
      }
    },
  )

  server.registerTool(
    'listLoopFilesInDrive',
    {
      description:
        'List .loop and .fluid files recursively inside a specific drive.',
      inputSchema: {
        driveId: z
          .string()
          .min(1)
          .max(512)
          .describe('Drive id returned by listLoopContainers.'),
        path: z
          .string()
          .max(1_024)
          .optional()
          .describe("Drive-relative folder path; default 'root'."),
        ...pageFields(DEFAULT_LIMITS.defaultFilePageSize),
      },
      annotations,
    },
    async args => {
      try {
        return toolJson(
          await service.listLoopFilesInDrive(
            args.driveId,
            args.path,
            pageRequest(args),
          ),
        )
      } catch (error) {
        return toolError(error)
      }
    },
  )

  return server
}
