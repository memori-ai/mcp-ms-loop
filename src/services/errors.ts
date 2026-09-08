import axios from 'axios'

export class LoopError extends Error {
  readonly code: string
  readonly status?: number

  constructor(message: string, code: string, status?: number) {
    super(message)
    this.name = 'LoopError'
    this.code = code
    this.status = status
  }
}

export function toLoopError(
  operation: string,
  error: unknown,
  fallbackCode = 'GRAPH_REQUEST_FAILED',
): LoopError {
  if (error instanceof LoopError) return error

  if (axios.isAxiosError(error)) {
    const status = error.response?.status

    if (status !== undefined) {
      return new LoopError(
        `${operation} failed with HTTP ${status}.`,
        fallbackCode,
        status,
      )
    }

    return new LoopError(
      `${operation} failed because the Microsoft service was unreachable.`,
      'NETWORK_ERROR',
    )
  }

  return new LoopError(`${operation} failed.`, fallbackCode)
}
