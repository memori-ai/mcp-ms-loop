import axios, { type AxiosResponse } from 'axios'

import {
  DEFAULT_LIMITS,
  jsonStringBytes,
  splitContentBudget,
  truncateText,
} from '../lib/limits.js'

import { getAccessToken, type LoopCredentials } from './authService.js'
import { LoopError, toLoopError } from './errors.js'
import type { HttpClient } from './http.js'

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0'
const SEARCH_REGIONS = ['EMEA', 'EUR', 'NAM', 'APAC', 'GCC', 'GCCHIGH'] as const
const SEARCH_PAGE_SIZE = 200
const SEARCH_MAX_HITS = 5_000
const REQUEST_TIMEOUT_MS = 30_000
const CHILDREN_SELECT = 'id,name,size,webUrl,lastModifiedDateTime,file,folder'

export type ContentMode = 'text' | 'html' | 'both'

export type ContentOptions = {
  contentMode?: ContentMode
  includeText?: boolean
  maxChars?: number
}

export type PageRequest = {
  offset: number
  limit: number
}

export type PageInfo = {
  offset: number
  limite: number
  restituiti: number
  haAltri: boolean
}

export type LoopContainer = {
  driveId: string
  name: string | null
  webUrl: string | null
  loopFileCount: number
}

export type LoopFile = {
  id: string
  name: string
  driveId: string
  size: number | null
  webUrl: string | null
  lastModified: string | null
  mimeType: string | null
}

type DriveItem = {
  id?: unknown
  name?: unknown
  size?: unknown
  webUrl?: unknown
  lastModifiedDateTime?: unknown
  file?: { mimeType?: unknown }
  folder?: unknown
  parentReference?: { driveId?: unknown }
}

type SearchHit = {
  resource?: DriveItem
}

type SearchResponse = {
  value?: Array<{
    hitsContainers?: Array<{
      total?: number
      hits?: SearchHit[]
    }>
  }>
}

type ChildrenResponse = {
  value?: DriveItem[]
  '@odata.nextLink'?: string
}

type LoopContentResult = {
  name: string | null
  driveId: string
  itemId: string
  webUrl: string | null
  lastModified: string | null
  size: number | null
  mimeType: string | null
  html?: string
  truncated?: boolean
  text?: string
  textTruncated?: boolean
}

export function encodeShareUrl(url: string): string {
  return `u!${Buffer.from(url, 'utf8')
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '')}`
}

export function isSpeUrl(shareUrl: string): boolean {
  try {
    const url = new URL(shareUrl)
    if (url.pathname.toLowerCase().includes('/contentstorage/csp_')) {
      return true
    }
    const navParam = url.searchParams.get('nav')
    if (!navParam) return false
    const storagePath = decodeSpeNavParams(navParam).get('s')
    return storagePath?.toLowerCase().includes('/contentstorage/csp_') ?? false
  } catch {
    return false
  }
}

export function parseSpeNavParam(navParam: string): {
  driveId: string | null
  itemId: string | null
} {
  try {
    const params = decodeSpeNavParams(navParam)
    return {
      driveId: params.get('d') || null,
      itemId: params.get('f') || null,
    }
  } catch {
    return { driveId: null, itemId: null }
  }
}

function decodeSpeNavParams(navParam: string): URLSearchParams {
  const normalized = navParam.replace(/-/g, '+').replace(/_/g, '/')
  const decoded = Buffer.from(normalized, 'base64').toString('utf8')
  return new URLSearchParams(decoded)
}

export function htmlToText(html: string): string {
  if (!html) return ''
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim()
}

export function resolveContentMode(options: ContentOptions): ContentMode {
  if (options.contentMode) return options.contentMode
  if (options.includeText === true) return 'both'
  if (options.includeText === false) return 'html'
  return 'text'
}

function asString(value: unknown): string | null {
  return typeof value === 'string' ? value : null
}

function asNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null
}

function asMimeType(item: DriveItem): string | null {
  return asString(item.file?.mimeType)
}

function graphHeaders(accessToken: string) {
  return { Authorization: `Bearer ${accessToken}` }
}

function validatedGraphNextLink(value: unknown): string | undefined {
  if (value === undefined) return undefined
  if (typeof value !== 'string') {
    throw new LoopError(
      'Graph returned an invalid nextLink.',
      'INVALID_NEXT_LINK',
    )
  }
  const url = new URL(value)
  if (
    url.protocol !== 'https:' ||
    url.hostname !== 'graph.microsoft.com' ||
    !url.pathname.startsWith('/v1.0/')
  ) {
    throw new LoopError(
      'Graph returned an invalid nextLink.',
      'INVALID_NEXT_LINK',
    )
  }
  return url.toString()
}

function pageInfo(
  request: PageRequest,
  returned: number,
  hasMore: boolean,
): PageInfo {
  return {
    offset: request.offset,
    limite: request.limit,
    restituiti: returned,
    haAltri: hasMore,
  }
}

export class LoopService {
  private cachedAccessToken: { value: string; expiresAt: number } | undefined
  private pendingAccessToken: Promise<string> | undefined

  constructor(
    private readonly credentials: LoopCredentials,
    private readonly maxContentChars: number,
    private readonly http: HttpClient = axios,
  ) {}

  async getLoopByShareUrl(
    shareUrl: string,
    options: ContentOptions = {},
  ): Promise<LoopContentResult> {
    try {
      const parsedUrl = new URL(shareUrl)
      if (!['http:', 'https:'].includes(parsedUrl.protocol)) {
        throw new LoopError(
          'shareUrl must use HTTP or HTTPS.',
          'INVALID_SHARE_URL',
        )
      }

      const accessToken = await this.accessToken()
      let driveId: string
      let itemId: string
      let metadata: DriveItem

      if (isSpeUrl(shareUrl)) {
        const navParam = parsedUrl.searchParams.get('nav')
        if (!navParam) {
          throw new LoopError(
            'The SharePoint Embedded URL has no nav parameter. Browse with listLoopContainers and listLoopFilesInDrive, then use getLoopByItemId.',
            'INVALID_SPE_URL',
          )
        }
        const ids = parseSpeNavParam(navParam)
        if (!ids.driveId || !ids.itemId) {
          throw new LoopError(
            'The SharePoint Embedded nav parameter has no usable driveId/itemId.',
            'INVALID_SPE_URL',
          )
        }
        driveId = ids.driveId
        itemId = ids.itemId
        metadata = await this.fetchDriveItemMetadata(
          accessToken,
          driveId,
          itemId,
        )
      } else {
        const response = await this.http.get<DriveItem>(
          `${GRAPH_BASE_URL}/shares/${encodeShareUrl(
            shareUrl,
          )}/driveItem?$select=${CHILDREN_SELECT},parentReference`,
          {
            headers: graphHeaders(accessToken),
            timeout: REQUEST_TIMEOUT_MS,
          },
        )
        metadata = response.data
        const resolvedDriveId = asString(metadata.parentReference?.driveId)
        const resolvedItemId = asString(metadata.id)
        if (!resolvedDriveId || !resolvedItemId) {
          throw new LoopError(
            'Microsoft Graph did not resolve driveId/itemId from the share URL.',
            'SHARE_NOT_RESOLVED',
          )
        }
        driveId = resolvedDriveId
        itemId = resolvedItemId
      }

      const html = await this.fetchLoopHtml(accessToken, driveId, itemId)
      return this.buildContentResult(metadata, driveId, itemId, html, options)
    } catch (error) {
      throw toLoopError('Reading the Loop sharing URL', error)
    }
  }

  async getLoopByItemId(
    driveId: string,
    itemId: string,
    options: ContentOptions = {},
  ): Promise<LoopContentResult> {
    try {
      const accessToken = await this.accessToken()
      const [metadata, html] = await Promise.all([
        this.fetchDriveItemMetadata(accessToken, driveId, itemId),
        this.fetchLoopHtml(accessToken, driveId, itemId),
      ])
      return this.buildContentResult(metadata, driveId, itemId, html, options)
    } catch (error) {
      throw toLoopError('Reading the Loop item', error)
    }
  }

  async listLoopContainers(request: PageRequest): Promise<{
    items: LoopContainer[]
    pagina: PageInfo & {
      totaleScoperti: number
      risultatiRicercaParziali: boolean
    }
  }> {
    try {
      const accessToken = await this.accessToken()
      const search = await this.searchLoopFiles(
        accessToken,
        request.offset + request.limit + 1,
      )
      const driveMap = new Map<string, LoopContainer>()

      for (const hit of search.hits) {
        const driveId = asString(hit.resource?.parentReference?.driveId)
        if (!driveId) continue
        const existing = driveMap.get(driveId)
        if (existing) {
          existing.loopFileCount += 1
        } else {
          driveMap.set(driveId, {
            driveId,
            name: null,
            webUrl: null,
            loopFileCount: 1,
          })
        }
      }

      const all = [...driveMap.values()]
      const items = all.slice(request.offset, request.offset + request.limit)
      await Promise.all(
        items.map(async workspace => {
          try {
            const response = await this.http.get<DriveItem>(
              `${GRAPH_BASE_URL}/drives/${encodeURIComponent(
                workspace.driveId,
              )}?$select=id,name,webUrl`,
              {
                headers: graphHeaders(accessToken),
                timeout: REQUEST_TIMEOUT_MS,
              },
            )
            workspace.name = asString(response.data.name) ?? workspace.driveId
            workspace.webUrl = asString(response.data.webUrl)
          } catch {
            workspace.name = workspace.driveId
          }
        }),
      )

      return {
        items,
        pagina: {
          ...pageInfo(
            request,
            items.length,
            request.offset + items.length < all.length ||
              search.moreRetrievable,
          ),
          totaleScoperti: all.length,
          risultatiRicercaParziali: search.partial,
        },
      }
    } catch (error) {
      throw toLoopError('Listing Loop workspaces', error)
    }
  }

  async listLoopFilesInDrive(
    driveId: string,
    path: string | undefined,
    request: PageRequest,
  ): Promise<{ items: LoopFile[]; pagina: PageInfo }> {
    try {
      const accessToken = await this.accessToken()
      const collected: LoopFile[] = []
      const visitedFolders = new Set<string>()
      let matchingIndex = 0

      const scanFolder = async (initialUrl: string): Promise<boolean> => {
        let nextLink: string | undefined = initialUrl
        while (nextLink) {
          const response: AxiosResponse<ChildrenResponse> =
            await this.http.get<ChildrenResponse>(nextLink, {
              headers: graphHeaders(accessToken),
              timeout: REQUEST_TIMEOUT_MS,
            })
          const items = response.data.value ?? []

          for (const item of items) {
            const id = asString(item.id)
            if (item.folder && id && !visitedFolders.has(id)) {
              visitedFolders.add(id)
              const stopped = await scanFolder(
                `${GRAPH_BASE_URL}/drives/${encodeURIComponent(
                  driveId,
                )}/items/${encodeURIComponent(
                  id,
                )}/children?$select=${CHILDREN_SELECT}`,
              )
              if (stopped) return true
            } else if (item.file) {
              const name = asString(item.name) ?? ''
              const lowerName = name.toLowerCase()
              if (lowerName.endsWith('.loop') || lowerName.endsWith('.fluid')) {
                if (matchingIndex >= request.offset && id) {
                  collected.push({
                    id,
                    name,
                    driveId,
                    size: asNumber(item.size),
                    webUrl: asString(item.webUrl),
                    lastModified: asString(item.lastModifiedDateTime),
                    mimeType: asMimeType(item),
                  })
                  if (collected.length > request.limit) return true
                }
                matchingIndex += 1
              }
            }
          }

          nextLink = validatedGraphNextLink(response.data['@odata.nextLink'])
        }
        return false
      }

      const rootUrl = this.childrenRootUrl(driveId, path)
      await scanFolder(rootUrl)
      const hasMore = collected.length > request.limit
      const items = collected.slice(0, request.limit)
      return {
        items,
        pagina: pageInfo(request, items.length, hasMore),
      }
    } catch (error) {
      throw toLoopError('Listing Loop files', error)
    }
  }

  private async accessToken(): Promise<string> {
    if (
      this.cachedAccessToken &&
      this.cachedAccessToken.expiresAt > Date.now()
    ) {
      return this.cachedAccessToken.value
    }
    if (this.pendingAccessToken) return this.pendingAccessToken

    const pending = getAccessToken(this.credentials, this.http).then(token => {
      const cacheLifetimeMs = (token.expiresInSeconds - 60) * 1_000
      if (cacheLifetimeMs > 0) {
        this.cachedAccessToken = {
          value: token.value,
          expiresAt: Date.now() + cacheLifetimeMs,
        }
      } else {
        this.cachedAccessToken = undefined
      }
      return token.value
    })
    this.pendingAccessToken = pending
    try {
      return await pending
    } finally {
      if (this.pendingAccessToken === pending) {
        this.pendingAccessToken = undefined
      }
    }
  }

  private buildContentResult(
    metadata: DriveItem,
    driveId: string,
    itemId: string,
    html: string,
    options: ContentOptions,
  ): LoopContentResult {
    const mode = resolveContentMode(options)
    const configuredMax = Math.max(
      DEFAULT_LIMITS.minimumContentChars,
      this.maxContentChars,
    )
    const maxChars = Math.min(
      Math.max(
        DEFAULT_LIMITS.minimumContentChars,
        options.maxChars ?? configuredMax,
      ),
      configuredMax,
    )
    const boundedMetadata = (
      value: string | null,
      budget: number,
    ): string | null =>
      value === null ? null : truncateText(value, budget).value
    const metadataBudgets =
      maxChars < 5_000
        ? {
            name: 96,
            id: 192,
            webUrl: 160,
            lastModified: 80,
            mimeType: 80,
          }
        : {
            name: 512,
            id: 512,
            webUrl: 1_500,
            lastModified: 192,
            mimeType: 256,
          }
    const result: LoopContentResult = {
      name: boundedMetadata(asString(metadata.name), metadataBudgets.name),
      driveId: boundedMetadata(driveId, metadataBudgets.id) ?? '',
      itemId: boundedMetadata(itemId, metadataBudgets.id) ?? '',
      webUrl: boundedMetadata(
        asString(metadata.webUrl),
        metadataBudgets.webUrl,
      ),
      lastModified: boundedMetadata(
        asString(metadata.lastModifiedDateTime),
        metadataBudgets.lastModified,
      ),
      size: asNumber(metadata.size),
      mimeType: boundedMetadata(asMimeType(metadata), metadataBudgets.mimeType),
    }

    if (mode === 'html') {
      const scaffold = { ...result, html: '', truncated: false }
      const fieldBudget = Math.max(
        2,
        maxChars - Buffer.byteLength(JSON.stringify(scaffold), 'utf8') + 2,
      )
      const bounded = truncateText(html, fieldBudget)
      result.html = bounded.value
      result.truncated = bounded.truncated
      return result
    }

    const text = htmlToText(html)
    if (mode === 'text') {
      const scaffold = { ...result, text: '', textTruncated: false }
      const fieldBudget = Math.max(
        2,
        maxChars - Buffer.byteLength(JSON.stringify(scaffold), 'utf8') + 2,
      )
      const bounded = truncateText(text, fieldBudget)
      result.text = bounded.value
      result.textTruncated = bounded.truncated
      return result
    }

    const scaffold = {
      ...result,
      html: '',
      truncated: false,
      text: '',
      textTruncated: false,
    }
    const totalFieldBudget = Math.max(
      4,
      maxChars - Buffer.byteLength(JSON.stringify(scaffold), 'utf8') + 4,
    )
    const budgets = splitContentBudget(
      jsonStringBytes(html),
      jsonStringBytes(text),
      totalFieldBudget,
    )
    const boundedHtml = truncateText(html, budgets.html)
    const boundedText = truncateText(text, budgets.text)
    result.html = boundedHtml.value
    result.truncated = boundedHtml.truncated
    result.text = boundedText.value
    result.textTruncated = boundedText.truncated
    return result
  }

  private async fetchLoopHtml(
    accessToken: string,
    driveId: string,
    itemId: string,
  ): Promise<string> {
    const itemPath = `${GRAPH_BASE_URL}/drives/${encodeURIComponent(
      driveId,
    )}/items/${encodeURIComponent(itemId)}/content`
    try {
      const response = await this.http.get<unknown>(`${itemPath}?format=html`, {
        headers: graphHeaders(accessToken),
        responseType: 'text',
        maxRedirects: 10,
        maxContentLength: Math.max(1_000_000, this.maxContentChars * 8),
        timeout: REQUEST_TIMEOUT_MS,
        transformResponse: [data => data],
      })
      return typeof response.data === 'string'
        ? response.data
        : Buffer.from(response.data as ArrayBuffer).toString('utf8')
    } catch (error) {
      if (
        axios.isAxiosError(error) &&
        [400, 415].includes(error.response?.status ?? 0)
      ) {
        const response = await this.http.get<ArrayBuffer>(itemPath, {
          headers: {
            ...graphHeaders(accessToken),
            Range: 'bytes=0-1499',
          },
          responseType: 'arraybuffer',
          maxRedirects: 10,
          maxContentLength: 4_096,
          timeout: REQUEST_TIMEOUT_MS,
        })
        const base64 = Buffer.from(response.data).toString('base64')
        const preview = base64.slice(0, 2_000)
        return `<!-- Loop raw content preview; HTML conversion unavailable. --><pre data-encoding="base64">${preview}${
          base64.length > preview.length ? '…[truncated]' : ''
        }</pre>`
      }
      throw error
    }
  }

  private async fetchDriveItemMetadata(
    accessToken: string,
    driveId: string,
    itemId: string,
  ): Promise<DriveItem> {
    const response = await this.http.get<DriveItem>(
      `${GRAPH_BASE_URL}/drives/${encodeURIComponent(
        driveId,
      )}/items/${encodeURIComponent(
        itemId,
      )}?$select=${CHILDREN_SELECT},parentReference`,
      {
        headers: graphHeaders(accessToken),
        timeout: REQUEST_TIMEOUT_MS,
      },
    )
    return response.data
  }

  private async searchLoopFiles(
    accessToken: string,
    minimumDistinctDrives: number,
  ): Promise<{
    hits: SearchHit[]
    partial: boolean
    moreRetrievable: boolean
  }> {
    const tried = new Set<string>()
    let lastRegionError: unknown
    let selectedRegion: string | undefined
    let firstPage: { hits: SearchHit[]; total: number | undefined } | undefined

    const searchRegion = async (region: string, from: number) => {
      tried.add(region)
      return this.http.post<SearchResponse>(
        `${GRAPH_BASE_URL}/search/query`,
        {
          requests: [
            {
              entityTypes: ['driveItem'],
              query: {
                queryString: 'filetype:loop OR filetype:fluid',
              },
              from,
              size: SEARCH_PAGE_SIZE,
              region,
            },
          ],
        },
        {
          headers: {
            ...graphHeaders(accessToken),
            'Content-Type': 'application/json',
          },
          timeout: REQUEST_TIMEOUT_MS,
        },
      )
    }

    for (const region of SEARCH_REGIONS) {
      if (tried.has(region)) continue
      try {
        const response = await searchRegion(region, 0)
        selectedRegion = region
        firstPage = this.parseSearchResponse(response.data)
        break
      } catch (error) {
        const requiredRegion = this.requiredSearchRegion(error)
        if (requiredRegion && !tried.has(requiredRegion)) {
          try {
            const response = await searchRegion(requiredRegion, 0)
            selectedRegion = requiredRegion
            firstPage = this.parseSearchResponse(response.data)
            break
          } catch (retryError) {
            if (!this.isSearchRegionError(retryError)) throw retryError
            lastRegionError = retryError
          }
        }
        if (this.isSearchRegionError(error)) {
          lastRegionError = error
          continue
        }
        throw error
      }
    }

    if (!selectedRegion || !firstPage) {
      if (lastRegionError) throw lastRegionError
      return { hits: [], partial: false, moreRetrievable: false }
    }

    const hits = [...firstPage.hits]
    const total = firstPage.total
    let exhausted =
      firstPage.hits.length < SEARCH_PAGE_SIZE ||
      (total !== undefined && hits.length >= total)

    while (
      !exhausted &&
      hits.length < SEARCH_MAX_HITS &&
      this.distinctDriveCount(hits) < minimumDistinctDrives
    ) {
      const response = await searchRegion(selectedRegion, hits.length)
      const page = this.parseSearchResponse(response.data)
      if (page.hits.length === 0) {
        exhausted = true
        break
      }
      hits.push(...page.hits.slice(0, SEARCH_MAX_HITS - hits.length))
      exhausted =
        page.hits.length < SEARCH_PAGE_SIZE ||
        (page.total !== undefined && hits.length >= page.total)
    }

    const capped = !exhausted && hits.length >= SEARCH_MAX_HITS
    return {
      hits,
      partial: !exhausted || (total !== undefined && hits.length < total),
      moreRetrievable: !exhausted && !capped,
    }
  }

  private parseSearchResponse(data: SearchResponse): {
    hits: SearchHit[]
    total: number | undefined
  } {
    const container = data.value?.[0]?.hitsContainers?.[0]
    return {
      hits: container?.hits ?? [],
      total: typeof container?.total === 'number' ? container.total : undefined,
    }
  }

  private distinctDriveCount(hits: SearchHit[]): number {
    return new Set(
      hits
        .map(hit => asString(hit.resource?.parentReference?.driveId))
        .filter((driveId): driveId is string => driveId !== null),
    ).size
  }

  private graphMessage(error: unknown): string {
    if (!axios.isAxiosError(error)) return ''
    const data = error.response?.data as
      | { error?: { message?: unknown } }
      | undefined
    return typeof data?.error?.message === 'string' ? data.error.message : ''
  }

  private isSearchRegionError(error: unknown): boolean {
    const message = this.graphMessage(error).toLowerCase()
    return (
      message.includes('only valid regions are') ||
      message.includes('region is required')
    )
  }

  private requiredSearchRegion(error: unknown): string | null {
    const message = this.graphMessage(error)
    const match = /Only valid regions are\s*:?\s*([A-Za-z]+)/i.exec(message)
    return match?.[1]?.toUpperCase() ?? null
  }

  private childrenRootUrl(driveId: string, path: string | undefined): string {
    const drivePath = `${GRAPH_BASE_URL}/drives/${encodeURIComponent(driveId)}`
    if (!path || path === '/' || path.toLowerCase() === 'root') {
      return `${drivePath}/root/children?$select=${CHILDREN_SELECT}`
    }

    const cleanPath = path.replace(/^\/|\/$/g, '')
    if (
      !cleanPath ||
      cleanPath.split('/').some(part => part === '..' || part === '.')
    ) {
      throw new LoopError(
        'path must be a drive-relative folder path.',
        'INVALID_PATH',
      )
    }
    const encodedPath = cleanPath
      .split('/')
      .map(part => encodeURIComponent(part))
      .join('/')
    return `${drivePath}/root:/${encodedPath}:/children?$select=${CHILDREN_SELECT}`
  }
}
