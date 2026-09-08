import { createEnv } from '@t3-oss/env-core'
import { z } from 'zod'

import { DEFAULT_LIMITS } from './limits.js'

export function zEnvInt(options: { min?: number; max?: number } = {}) {
  let schema = z.coerce.number().int()
  if (options.min !== undefined) schema = schema.min(options.min)
  if (options.max !== undefined) schema = schema.max(options.max)
  return schema
}

export function loadEnv(runtimeEnv: NodeJS.ProcessEnv = process.env) {
  return createEnv({
    server: {
      TENANT_ID: z.string().min(1, 'required'),
      CLIENT_ID: z.string().min(1, 'required'),
      CLIENT_SECRET: z.string().min(1, 'required'),
      LOOP_MAX_CONTENT_CHARS: zEnvInt({
        min: DEFAULT_LIMITS.minimumContentChars,
        max: DEFAULT_LIMITS.absoluteMaxContentChars,
      }).default(DEFAULT_LIMITS.maxContentChars),
    },
    runtimeEnv,
    emptyStringAsUndefined: true,
    onValidationError: issues => {
      const detail = issues
        .map(issue => {
          const path =
            (issue.path ?? [])
              .map(part =>
                typeof part === 'string' || typeof part === 'number'
                  ? String(part)
                  : JSON.stringify(part),
              )
              .join('.') || '(root)'
          return `  - ${path}: ${issue.message}`
        })
        .join('\n')
      throw new Error(`Invalid environment:\n${detail}`)
    },
  })
}

export type LoopEnv = ReturnType<typeof loadEnv>
