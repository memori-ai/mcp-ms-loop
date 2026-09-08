import type { AxiosRequestConfig, AxiosResponse } from 'axios'
import { describe, expect, it } from 'vitest'

import type { HttpClient } from '../src/services/http.js'
import {
  LoopService,
  encodeShareUrl,
  isSpeUrl,
  parseSpeNavParam,
} from '../src/services/loopService.js'

type Step = {
  method: 'GET' | 'POST'
  urlIncludes: string
  data?: unknown
  error?: unknown
}

class FakeHttp implements HttpClient {
  readonly calls: Array<{
    method: string
    url: string
    data?: unknown
    config?: AxiosRequestConfig
  }> = []

  constructor(private readonly steps: Step[]) {}

  async get<T = unknown>(
    url: string,
    config?: AxiosRequestConfig,
  ): Promise<AxiosResponse<T>> {
    return this.take<T>('GET', url, undefined, config)
  }

  async post<T = unknown>(
    url: string,
    data?: unknown,
    config?: AxiosRequestConfig,
  ): Promise<AxiosResponse<T>> {
    return this.take<T>('POST', url, data, config)
  }

  expectDone() {
    expect(this.steps).toHaveLength(0)
  }

  private async take<T>(
    method: 'GET' | 'POST',
    url: string,
    data?: unknown,
    config?: AxiosRequestConfig,
  ): Promise<AxiosResponse<T>> {
    this.calls.push({ method, url, data, config })
    const step = this.steps.shift()
    expect(step, `unexpected ${method} ${url}`).toBeDefined()
    expect(method).toBe(step?.method)
    expect(url).toContain(step?.urlIncludes)
    if (step?.error) throw step.error
    return {
      data: step?.data as T,
      status: 200,
      statusText: 'OK',
      headers: {},
      config: { headers: {} },
    } as AxiosResponse<T>
  }
}

const credentials = {
  tenantId: 'tenant-id',
  clientId: 'client-id',
  clientSecret: 'private-client-secret',
}

const tokenStep: Step = {
  method: 'POST',
  urlIncludes: '/tenant-id/oauth2/v2.0/token',
  data: { access_token: 'access-token' },
}

const metadata = {
  id: 'item-id',
  name: 'Example.loop',
  size: 123,
  webUrl: 'https://example.invalid/file',
  lastModifiedDateTime: '2026-09-08T10:00:00Z',
  file: { mimeType: 'application/octet-stream' },
  parentReference: { driveId: 'drive-id' },
}

function axiosLikeError(status: number, code: string, message: string) {
  return {
    isAxiosError: true,
    message,
    response: {
      status,
      data: { error: { code, message } },
    },
  }
}

describe('LoopService content reads', () => {
  it('resolves a standard share and returns text only by default', async () => {
    const shareUrl = 'https://example.invalid/shared/file'
    const http = new FakeHttp([
      tokenStep,
      {
        method: 'GET',
        urlIncludes: `/shares/${encodeShareUrl(shareUrl)}/driveItem`,
        data: metadata,
      },
      {
        method: 'GET',
        urlIncludes: '/drives/drive-id/items/item-id/content?format=html',
        data: '<p>Hello &amp; world</p>',
      },
    ])
    const service = new LoopService(credentials, 1_000, http)

    const result = await service.getLoopByShareUrl(shareUrl)

    expect(result.text).toBe('Hello & world')
    expect(result.textTruncated).toBe(false)
    expect(result).not.toHaveProperty('html')
    http.expectDone()
  })

  it('keeps explicit legacy includeText=true within one total budget', async () => {
    const html = `<p>${'x'.repeat(10_000)}</p>`
    const http = new FakeHttp([
      tokenStep,
      {
        method: 'GET',
        urlIncludes: '/drives/drive-id/items/item-id?$select=',
        data: metadata,
      },
      {
        method: 'GET',
        urlIncludes: '/drives/drive-id/items/item-id/content?format=html',
        data: html,
      },
    ])
    const service = new LoopService(credentials, 5_000, http)

    const result = await service.getLoopByItemId('drive-id', 'item-id', {
      includeText: true,
    })

    expect(result.html).toBeDefined()
    expect(result.text).toBeDefined()
    expect(Buffer.byteLength(JSON.stringify(result))).toBeLessThanOrEqual(5_000)
    expect(result.truncated).toBe(true)
    expect(result.textTruncated).toBe(true)
    http.expectDone()
  })

  it('bounds the complete serialized response after JSON escaping', async () => {
    const http = new FakeHttp([
      tokenStep,
      {
        method: 'GET',
        urlIncludes: '/drives/drive-id/items/item-id?$select=',
        data: metadata,
      },
      {
        method: 'GET',
        urlIncludes: '/content?format=html',
        data: `<p>${'\0"\\😀'.repeat(2_000)}</p>`,
      },
    ])
    const service = new LoopService(credentials, 1_000, http)

    const result = await service.getLoopByItemId('drive-id', 'item-id')

    expect(Buffer.byteLength(JSON.stringify(result))).toBeLessThanOrEqual(1_000)
    expect(result.textTruncated).toBe(true)
    http.expectDone()
  })

  it('parses SharePoint Embedded nav links without calling /shares', async () => {
    const nav = Buffer.from(
      's=%2Fcontentstorage%2FCSP_example&d=drive-spe&f=item-spe',
    ).toString('base64')
    const shareUrl = `https://loop.cloud.microsoft/open?nav=${encodeURIComponent(
      nav,
    )}`
    const http = new FakeHttp([
      tokenStep,
      {
        method: 'GET',
        urlIncludes: '/drives/drive-spe/items/item-spe?$select=',
        data: { ...metadata, id: 'item-spe' },
      },
      {
        method: 'GET',
        urlIncludes: '/drives/drive-spe/items/item-spe/content?format=html',
        data: '<p>SPE content</p>',
      },
    ])
    const service = new LoopService(credentials, 5_000, http)

    const result = await service.getLoopByShareUrl(shareUrl)

    expect(parseSpeNavParam(nav)).toEqual({
      driveId: 'drive-spe',
      itemId: 'item-spe',
    })
    expect(isSpeUrl(shareUrl)).toBe(true)
    expect(
      isSpeUrl('https://example.invalid/open?next=/contentstorage/CSP_false'),
    ).toBe(false)
    expect(result.driveId).toBe('drive-spe')
    expect(http.calls.some(call => call.url.includes('/shares/'))).toBe(false)
    http.expectDone()
  })

  it('retains the bounded raw preview fallback for unsupported HTML conversion', async () => {
    const http = new FakeHttp([
      tokenStep,
      {
        method: 'GET',
        urlIncludes: '/drives/drive-id/items/item-id?$select=',
        data: metadata,
      },
      {
        method: 'GET',
        urlIncludes: '/content?format=html',
        error: axiosLikeError(415, 'unsupportedFormat', 'unsupported'),
      },
      {
        method: 'GET',
        urlIncludes: '/drives/drive-id/items/item-id/content',
        data: Uint8Array.from([1, 2, 3]).buffer,
      },
    ])
    const service = new LoopService(credentials, 5_000, http)

    const result = await service.getLoopByItemId('drive-id', 'item-id', {
      contentMode: 'html',
    })

    expect(result.html).toContain('Loop raw content preview')
    const rawCall = http.calls.find(
      call =>
        call.method === 'GET' &&
        call.url.endsWith('/drives/drive-id/items/item-id/content'),
    )
    expect(rawCall?.config?.headers).toMatchObject({
      Range: 'bytes=0-1499',
    })
    expect(rawCall?.config?.maxContentLength).toBe(4_096)
    expect(rawCall?.config?.timeout).toBe(30_000)
    http.expectDone()
  })

  it('does not cache short-lived OAuth tokens and reuses safe ones', async () => {
    const itemSteps: Step[] = [
      {
        method: 'GET',
        urlIncludes: '/items/item-id?$select=',
        data: metadata,
      },
      {
        method: 'GET',
        urlIncludes: '/content?format=html',
        data: '<p>content</p>',
      },
    ]
    const http = new FakeHttp([
      {
        ...tokenStep,
        data: { access_token: 'short-token', expires_in: 30 },
      },
      ...itemSteps,
      {
        ...tokenStep,
        data: { access_token: 'cacheable-token', expires_in: 120 },
      },
      ...itemSteps,
      ...itemSteps,
    ])
    const service = new LoopService(credentials, 5_000, http)

    await service.getLoopByItemId('drive-id', 'item-id')
    await service.getLoopByItemId('drive-id', 'item-id')
    await service.getLoopByItemId('drive-id', 'item-id')

    expect(
      http.calls.filter(call => call.url.includes('/oauth2/v2.0/token')),
    ).toHaveLength(2)
    http.expectDone()
  })
})

describe('LoopService bounded listings', () => {
  it('retries the tenant search region and paginates workspaces', async () => {
    const http = new FakeHttp([
      tokenStep,
      {
        method: 'POST',
        urlIncludes: '/search/query',
        error: axiosLikeError(
          400,
          'invalidRequest',
          'Only valid regions are EUR.',
        ),
      },
      {
        method: 'POST',
        urlIncludes: '/search/query',
        data: {
          value: [
            {
              hitsContainers: [
                {
                  total: 600,
                  hits: [
                    {
                      resource: {
                        parentReference: { driveId: 'drive-a' },
                      },
                    },
                    {
                      resource: {
                        parentReference: { driveId: 'drive-a' },
                      },
                    },
                    {
                      resource: {
                        parentReference: { driveId: 'drive-b' },
                      },
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
      {
        method: 'GET',
        urlIncludes: '/drives/drive-a?$select=id,name,webUrl',
        data: {
          name: 'Workspace A',
          webUrl: 'https://example.invalid/a',
        },
      },
    ])
    const service = new LoopService(credentials, 5_000, http)

    const result = await service.listLoopContainers({
      offset: 0,
      limit: 1,
    })

    expect(result.items).toEqual([
      {
        driveId: 'drive-a',
        name: 'Workspace A',
        webUrl: 'https://example.invalid/a',
        loopFileCount: 2,
      },
    ])
    expect(result.pagina).toMatchObject({
      restituiti: 1,
      haAltri: true,
      totaleScoperti: 2,
      risultatiRicercaParziali: true,
    })
    const retryBody = http.calls[2]?.data as {
      requests?: Array<{ region?: string }>
    }
    expect(retryBody.requests?.[0]?.region).toBe('EUR')
    http.expectDone()
  })

  it('continues Graph Search beyond the first hit page for later workspaces', async () => {
    const repeatedHits = Array.from({ length: 200 }, (_, index) => ({
      hitId: `hit-${index}`,
      resource: {
        parentReference: { driveId: 'drive-a' },
      },
    }))
    const http = new FakeHttp([
      tokenStep,
      {
        method: 'POST',
        urlIncludes: '/search/query',
        data: {
          value: [
            {
              hitsContainers: [{ total: 201, hits: repeatedHits }],
            },
          ],
        },
      },
      {
        method: 'POST',
        urlIncludes: '/search/query',
        data: {
          value: [
            {
              hitsContainers: [
                {
                  total: 201,
                  hits: [
                    {
                      resource: {
                        parentReference: { driveId: 'drive-b' },
                      },
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
      {
        method: 'GET',
        urlIncludes: '/drives/drive-b?$select=id,name,webUrl',
        data: {
          name: 'Workspace B',
          webUrl: 'https://example.invalid/b',
        },
      },
    ])
    const service = new LoopService(credentials, 5_000, http)

    const result = await service.listLoopContainers({
      offset: 1,
      limit: 1,
    })

    expect(result.items).toEqual([
      {
        driveId: 'drive-b',
        name: 'Workspace B',
        webUrl: 'https://example.invalid/b',
        loopFileCount: 1,
      },
    ])
    const secondSearch = http.calls[2]?.data as {
      requests?: Array<{ from?: number }>
    }
    expect(secondSearch.requests?.[0]?.from).toBe(200)
    expect(result.pagina.haAltri).toBe(false)
    http.expectDone()
  })

  it('stops recursive file scanning after page size plus one', async () => {
    const file = (id: string) => ({
      id,
      name: `${id}.loop`,
      size: 10,
      webUrl: null,
      lastModifiedDateTime: null,
      file: { mimeType: 'application/octet-stream' },
    })
    const http = new FakeHttp([
      tokenStep,
      {
        method: 'GET',
        urlIncludes: '/root/children?$select=',
        data: {
          value: [file('one'), { id: 'folder', folder: {} }],
        },
      },
      {
        method: 'GET',
        urlIncludes: '/items/folder/children?$select=',
        data: { value: [file('two'), file('three')] },
      },
    ])
    const service = new LoopService(credentials, 5_000, http)

    const result = await service.listLoopFilesInDrive('drive-id', 'root', {
      offset: 0,
      limit: 2,
    })

    expect(result.items.map(item => item.id)).toEqual(['one', 'two'])
    expect(result.pagina.haAltri).toBe(true)
    http.expectDone()
  })

  it('follows Graph nextLink while keeping a compact page', async () => {
    const http = new FakeHttp([
      tokenStep,
      {
        method: 'GET',
        urlIncludes: '/root/children?$select=',
        data: {
          value: [],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next-page',
        },
      },
      {
        method: 'GET',
        urlIncludes: '/v1.0/next-page',
        data: {
          value: [
            {
              id: 'found',
              name: 'Found.fluid',
              file: { mimeType: 'application/octet-stream' },
            },
          ],
        },
      },
    ])
    const service = new LoopService(credentials, 5_000, http)

    const result = await service.listLoopFilesInDrive('drive-id', undefined, {
      offset: 0,
      limit: 10,
    })

    expect(result.items.map(item => item.id)).toEqual(['found'])
    expect(result.pagina.haAltri).toBe(false)
    http.expectDone()
  })

  it('rejects off-origin Graph nextLink values before sending credentials', async () => {
    const http = new FakeHttp([
      tokenStep,
      {
        method: 'GET',
        urlIncludes: '/root/children?$select=',
        data: {
          value: [],
          '@odata.nextLink': 'https://attacker.invalid/collect',
        },
      },
    ])
    const service = new LoopService(credentials, 5_000, http)

    await expect(
      service.listLoopFilesInDrive('drive-id', undefined, {
        offset: 0,
        limit: 10,
      }),
    ).rejects.toMatchObject({ code: 'INVALID_NEXT_LINK' })
    expect(http.calls.some(call => call.url.includes('attacker.invalid'))).toBe(
      false,
    )
    http.expectDone()
  })
})

describe('LoopService error privacy', () => {
  it('never exposes OAuth request data or client secrets', async () => {
    const http = new FakeHttp([
      {
        method: 'POST',
        urlIncludes: '/oauth2/v2.0/token',
        error: axiosLikeError(
          401,
          'invalid_client',
          `Rejected ${credentials.clientSecret}`,
        ),
      },
    ])
    const service = new LoopService(credentials, 5_000, http)

    let thrown: unknown
    try {
      await service.getLoopByItemId('drive-id', 'item-id')
    } catch (error) {
      thrown = error
    }
    expect(thrown).toMatchObject({
      code: 'AUTHENTICATION_FAILED',
      status: 401,
    })
    expect(String(thrown)).not.toContain(credentials.clientSecret)
    http.expectDone()
  })
})
