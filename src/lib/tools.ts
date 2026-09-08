import { LoopError } from '../services/errors.js'

import { DEFAULT_LIMITS } from './limits.js'

export function errorText(error: unknown): string {
  const raw =
    error instanceof LoopError ? error.message : 'Unexpected server error.'
  return raw.slice(0, DEFAULT_LIMITS.maxErrorChars)
}

export function toolText(text: string) {
  return { content: [{ type: 'text' as const, text }] }
}

export function toolJson(value: unknown) {
  return toolText(JSON.stringify(value))
}

export function toolError(error: unknown) {
  const candidate = error instanceof LoopError ? error : undefined
  return {
    isError: true as const,
    content: [
      {
        type: 'text' as const,
        text: JSON.stringify({
          error: true,
          code: candidate?.code ?? 'UNEXPECTED',
          message: errorText(error),
          ...(candidate?.status === undefined
            ? {}
            : { status: candidate.status }),
        }),
      },
    ],
  }
}
