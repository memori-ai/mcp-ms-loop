export const DEFAULT_LIMITS = {
  minimumContentChars: 1_000,
  maxContentChars: 100_000,
  absoluteMaxContentChars: 400_000,
  defaultContainerPageSize: 20,
  defaultFilePageSize: 50,
  maxPageSize: 200,
  maxErrorChars: 500,
} as const

const TRUNCATION_MARKER = '\n…[content truncated]'

export function truncateText(
  text: string,
  maxBytes: number,
): { value: string; truncated: boolean } {
  if (jsonStringBytes(text) <= maxBytes) {
    return { value: text, truncated: false }
  }

  let low = 0
  let high = text.length
  let value = ''
  while (low <= high) {
    const middle = Math.floor((low + high) / 2)
    const candidate = `${text.slice(0, middle)}${TRUNCATION_MARKER}`
    if (jsonStringBytes(candidate) <= maxBytes) {
      value = candidate
      low = middle + 1
    } else {
      high = middle - 1
    }
  }

  return { value, truncated: true }
}

export function jsonStringBytes(text: string): number {
  return Buffer.byteLength(JSON.stringify(text), 'utf8')
}

export function splitContentBudget(
  htmlLength: number,
  textLength: number,
  totalBudget: number,
): { html: number; text: number } {
  let html = Math.floor(totalBudget / 2)
  let text = totalBudget - html

  if (htmlLength < html) {
    text += html - htmlLength
    html = htmlLength
  }
  if (textLength < text) {
    html += text - textLength
    text = textLength
  }

  return { html, text }
}
