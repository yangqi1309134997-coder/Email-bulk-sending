export const parseApiDate = (value) => {
  if (!value) return null
  if (value instanceof Date) return value
  const text = String(value)
  const hasTimezone = /(?:Z|[+-]\d{2}:?\d{2})$/i.test(text)
  const parsed = new Date(hasTimezone ? text : `${text}Z`)
  return Number.isNaN(parsed.getTime()) ? null : parsed
}

export const formatApiDateTime = (value, locale = 'zh-CN') => {
  const parsed = parseApiDate(value)
  return parsed ? parsed.toLocaleString(locale) : '-'
}
