/**
 * Display-only decoding of a bounded, logged Excel selection reference.
 * @param raw - serialized context JSON from the canonical user-message wrapper.
 * @returns the captured selection address, or undefined for ordinary or invalid text.
 */
export function excelContextLabel(raw: string): string | undefined {
  if (new TextEncoder().encode(raw).byteLength > 2048) return undefined
  let value: unknown
  try { value = JSON.parse(raw) } catch (_error) { return undefined /* Ordinary user text is not a context reference. */ }
  const record = (input: unknown): input is Record<string, unknown> =>
    typeof input === 'object' && input !== null && !Array.isArray(input)
  const text = (input: unknown): input is string => typeof input === 'string' && input.length > 0
  const keys = (input: Record<string, unknown>, allowed: readonly string[]): boolean =>
    Object.keys(input).every(key => allowed.includes(key))
  if (!record(value) || !record(value.workbook) || !record(value.worksheet) || !record(value.selection)) return undefined
  if (!keys(value, ['workbook', 'worksheet', 'selection', 'observedAt'])
    || !keys(value.workbook, ['id', 'name', 'path']) || !keys(value.worksheet, ['id', 'name'])
    || !keys(value.selection, ['workbookId', 'address'])
    || (value.workbook.name !== undefined && !text(value.workbook.name))
    || (value.workbook.path !== undefined && !text(value.workbook.path))) return undefined
  if (!text(value.workbook.id) || !text(value.worksheet.id) || !text(value.worksheet.name)
    || !text(value.selection.address) || value.selection.workbookId !== value.workbook.id
    || !text(value.observedAt) || !Number.isFinite(Date.parse(value.observedAt))) return undefined
  return value.selection.address
}
