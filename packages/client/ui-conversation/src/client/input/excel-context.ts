import { z } from 'zod'
import { randomUUID } from '@deepseek-ai/dsh-util-crypto'
import type { ExcelContextConfig } from '../../excel-config.ts'

export type { ExcelContextConfig } from '../../excel-config.ts'

const CITATION_INSTRUCTION = 'Cite Excel cells as [[cite:Sheet1!A1:B2]] (quote worksheet names with spaces).'

const text = z.string().min(1)
const contextSchema = z.strictObject({
  workbook: z.strictObject({ id: text, name: text.optional(), path: text.optional() }),
  worksheet: z.strictObject({ id: text, name: text }),
  selection: z.strictObject({ workbookId: text, address: text }),
  observedAt: z.iso.datetime(),
}).refine(context => context.workbook.id === context.selection.workbookId)

/**
 * Request one bounded snapshot from the configured Excel parent. Each call
 * starts immediately and owns its listener, timeout, correlation, and abort.
 * @param config - exact trusted parent origin and capture timeout.
 * @param signal - the submitting message's lifetime.
 * @param unavailable - localized capture failure notice.
 * @param target - browser document; tests provide a structural window.
 * @returns durable user-message text, or no context for a standalone document.
 */
export function captureExcelContext(
  config: ExcelContextConfig,
  signal: AbortSignal,
  unavailable: string,
  target: Window = window,
): Promise<string> | undefined {
  if (target.parent === target) return undefined
  const requestId = randomUUID()
  return new Promise<string>((resolve, reject) => {
    const finish = (result: string | Error): void => {
      target.removeEventListener('message', receive)
      signal.removeEventListener('abort', abort)
      clearTimeout(timeout)
      if (result instanceof Error) reject(result)
      else resolve(result)
    }
    const abort = (): void => { finish(new Error(unavailable)) }
    const receive = (event: MessageEvent<unknown>): void => {
      if (event.source !== target.parent || event.origin !== config.parentOrigin) return
      const envelope = z.strictObject({
        type: z.literal('dsh/excel-context/response'), version: z.literal(1), requestId: z.literal(requestId),
        context: z.unknown().optional(),
        error: z.enum(['context_unavailable', 'context_too_large']).optional(),
      }).safeParse(event.data)
      if (!envelope.success) return
      if (envelope.data.error !== undefined || envelope.data.context === undefined) {
        finish(new Error(unavailable))
        return
      }
      let raw: string | undefined
      try { raw = JSON.stringify(envelope.data.context) } catch (_error) {
        finish(new Error(unavailable))
        return
      }
      if (new TextEncoder().encode(raw).byteLength > 2048) {
        finish(new Error(unavailable))
        return
      }
      const context = contextSchema.safeParse(envelope.data.context)
      if (!context.success) {
        finish(new Error(unavailable))
        return
      }
      finish(`[Excel context]\n${JSON.stringify(context.data)}\n${CITATION_INSTRUCTION}\n[/Excel context]`)
    }
    const timeout = setTimeout(abort, config.timeoutMs)
    target.addEventListener('message', receive)
    signal.addEventListener('abort', abort, { once: true })
    if (signal.aborted) { abort(); return }
    try {
      target.parent.postMessage({ type: 'dsh/excel-context/request', version: 1, requestId }, config.parentOrigin)
    } catch (_error) { finish(new Error(unavailable)) }
  })
}

/**
 * Decode a bounded canonical logged marker without reading workbook contents.
 * @param message - text of one durable user message.
 * @returns its workbook binding, or absence for invalid/uncaptured input.
 */
export function excelWorkbookOf(message: string): string | undefined {
  const suffix = `\n${CITATION_INSTRUCTION}\n[/Excel context]`
  if (!message.endsWith(suffix)) return undefined
  const marker = /\[Excel context\]\n([^\n]*)$/u.exec(message.slice(0, -suffix.length))
  if (marker === null || new TextEncoder().encode(marker[1]).byteLength > 2048) return undefined
  let raw: unknown
  try { raw = JSON.parse(marker[1] ?? '') } catch (_error) { return undefined /* Invalid logged marker has no navigation authority. */ }
  const context = contextSchema.safeParse(raw)
  return context.success ? context.data.workbook.id : undefined
}

/**
 * Navigate through the same authenticated parent with a bounded operation lifetime.
 * @param config - trusted parent origin and timeout.
 * @param workbookId - originating user message's immutable workbook binding.
 * @param address - validated sheet-qualified cell reference.
 * @param signal - navigation lifetime owned by the client plugin.
 * @returns completion after a matching authenticated acknowledgement; rejects on refusal/timeout.
 */
export function navigateExcelReference(
  config: ExcelContextConfig, workbookId: string, address: string, signal: AbortSignal,
): Promise<void> {
  const requestId = randomUUID()
  return new Promise<void>((resolve, reject) => {
    const finish = (ok: boolean): void => {
      window.removeEventListener('message', receive)
      signal.removeEventListener('abort', abort)
      clearTimeout(timeout)
      if (ok) resolve()
      else reject(new Error('Excel cell reference could not be opened'))
    }
    const abort = (): void => { finish(false) }
    const receive = (event: MessageEvent<unknown>): void => {
      if (event.source !== window.parent || event.origin !== config.parentOrigin) return
      const reply = z.strictObject({ type: z.literal('dsh/excel-context/navigation'), version: z.literal(1),
        requestId: z.literal(requestId), ok: z.literal(true).optional(), error: z.literal('context_unavailable').optional() }).safeParse(event.data)
      if (reply.success) finish(reply.data.ok === true && reply.data.error === undefined)
    }
    const timeout = setTimeout(abort, config.timeoutMs)
    window.addEventListener('message', receive)
    signal.addEventListener('abort', abort, { once: true })
    if (signal.aborted) { abort(); return }
    try { window.parent.postMessage({ type: 'dsh/excel-context/navigate', version: 1, requestId, workbookId, address }, config.parentOrigin) }
    catch (_error) { finish(false) }
  })
}
