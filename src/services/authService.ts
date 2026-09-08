import axios from 'axios'

import { toLoopError } from './errors.js'
import type { HttpClient } from './http.js'

export type LoopCredentials = {
  tenantId: string
  clientId: string
  clientSecret: string
}

type TokenResponse = {
  access_token?: unknown
  expires_in?: unknown
}

export async function getAccessToken(
  credentials: LoopCredentials,
  http: HttpClient = axios,
): Promise<{ value: string; expiresInSeconds: number }> {
  const url = `https://login.microsoftonline.com/${encodeURIComponent(
    credentials.tenantId,
  )}/oauth2/v2.0/token`
  const params = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: credentials.clientId,
    client_secret: credentials.clientSecret,
    scope: 'https://graph.microsoft.com/.default',
  })

  try {
    const response = await http.post<TokenResponse>(url, params, {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      timeout: 30_000,
    })
    if (typeof response.data.access_token !== 'string') {
      throw new Error('Missing access token')
    }
    return {
      value: response.data.access_token,
      expiresInSeconds:
        typeof response.data.expires_in === 'number' &&
        Number.isFinite(response.data.expires_in) &&
        response.data.expires_in > 0
          ? response.data.expires_in
          : 3_600,
    }
  } catch (error) {
    throw toLoopError(
      'Microsoft identity authentication',
      error,
      'AUTHENTICATION_FAILED',
    )
  }
}
