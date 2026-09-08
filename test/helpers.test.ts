import { describe, expect, it } from 'vitest'

import { loadEnv, zEnvInt } from '../src/lib/env.js'
import {
  jsonStringBytes,
  splitContentBudget,
  truncateText,
} from '../src/lib/limits.js'
import { toolError, toolJson } from '../src/lib/tools.js'
import { LoopError } from '../src/services/errors.js'
import { resolveContentMode } from '../src/services/loopService.js'

describe('environment', () => {
  it('validates the existing Gateway variable names', () => {
    const env = loadEnv({
      TENANT_ID: 'tenant',
      CLIENT_ID: 'client',
      CLIENT_SECRET: 'secret',
      LOOP_MAX_CONTENT_CHARS: '120000',
    })
    expect(env.LOOP_MAX_CONTENT_CHARS).toBe(120_000)
  })

  it('bounds numeric environment values', () => {
    const schema = zEnvInt({ min: 1, max: 10 })
    expect(schema.parse('5')).toBe(5)
    expect(() => schema.parse('0')).toThrow()
    expect(() => schema.parse('11')).toThrow()
  })
})

describe('token limits', () => {
  it('keeps truncation including its marker inside the budget', () => {
    const result = truncateText('x'.repeat(2_000), 1_000)
    expect(result.truncated).toBe(true)
    expect(jsonStringBytes(result.value)).toBeLessThanOrEqual(1_000)
  })

  it('counts JSON escaping and UTF-8 bytes inside the budget', () => {
    const result = truncateText('\0"\\😀'.repeat(1_000), 1_000)
    expect(result.truncated).toBe(true)
    expect(jsonStringBytes(result.value)).toBeLessThanOrEqual(1_000)
  })

  it('shares and reallocates the total content budget', () => {
    expect(splitContentBudget(2_000, 2_000, 1_000)).toEqual({
      html: 500,
      text: 500,
    })
    expect(splitContentBudget(100, 2_000, 1_000)).toEqual({
      html: 100,
      text: 900,
    })
  })
})

describe('content mode compatibility', () => {
  it('uses compact text by default and preserves explicit legacy flags', () => {
    expect(resolveContentMode({})).toBe('text')
    expect(resolveContentMode({ includeText: true })).toBe('both')
    expect(resolveContentMode({ includeText: false })).toBe('html')
    expect(
      resolveContentMode({
        contentMode: 'text',
        includeText: true,
      }),
    ).toBe('text')
  })
})

describe('MCP envelopes', () => {
  it('uses compact JSON and structured errors', () => {
    expect(toolJson({ ok: true }).content[0]?.text).toBe('{"ok":true}')
    const result = toolError(
      new LoopError('Request denied.', 'ACCESS_DENIED', 403),
    )
    expect(result.isError).toBe(true)
    expect(JSON.parse(result.content[0]?.text ?? '{}')).toEqual({
      error: true,
      code: 'ACCESS_DENIED',
      message: 'Request denied.',
      status: 403,
    })

    const unexpected = JSON.stringify(
      toolError(new Error('private-client-secret')),
    )
    expect(unexpected).not.toContain('private-client-secret')
    expect(unexpected).toContain('Unexpected server error.')
  })
})
