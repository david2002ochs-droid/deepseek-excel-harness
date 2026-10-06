// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { captureExcelContext, excelWorkbookOf, navigateExcelReference } from '../src/client/input/excel-context.ts'

const config = { parentOrigin: 'https://localhost:3443', timeoutMs: 1000 }
const context = {
  workbook: { id: 'book-1', name: 'Budget.xlsx', path: 'C:\\Books\\Budget.xlsx' },
  worksheet: { id: 'sheet-1', name: 'Forecast' },
  selection: { workbookId: 'book-1', address: "'Forecast'!A1:B4" },
  observedAt: '2026-10-06T12:00:00.000Z',
}

const frames = new Set<HTMLIFrameElement>()

function embedding() {
  const frame = document.createElement('iframe')
  frames.add(frame)
  document.body.appendChild(frame)
  const parent = frame.contentWindow!
  const postMessage = vi.fn<(message: { requestId: unknown }, origin: string) => void>()
  vi.spyOn(parent, 'postMessage').mockImplementation((message: unknown, origin?: string | WindowPostMessageOptions) => {
    if (typeof message !== 'object' || message === null || !('requestId' in message)
      || typeof message.requestId !== 'string' || typeof origin !== 'string') {
      throw new Error('Expected a correlated Excel request with an explicit origin')
    }
    postMessage(message, origin)
  })
  vi.spyOn(window, 'parent', 'get').mockReturnValue(parent)
  const reply = (data: unknown, origin = config.parentOrigin, source: MessageEventSource = parent) => {
    window.dispatchEvent(new MessageEvent('message', { data, origin, source }))
  }
  const envelope = (payload: unknown = context) => ({
    type: 'dsh/excel-context/response', version: 1,
    requestId: postMessage.mock.calls.at(-1)?.[0].requestId, context: payload,
  })
  return { postMessage, parent, reply, envelope }
}

afterEach(() => {
  vi.restoreAllMocks()
  vi.useRealTimers()
  for (const frame of frames) frame.remove()
  frames.clear()
})

describe('Excel context capture', () => {
  it('leaves an ordinary standalone Harness document unchanged', () => {
    expect(captureExcelContext(config, new AbortController().signal, 'unavailable')).toBeUndefined()
  })

  it('accepts only the configured origin, exact parent source, and matching request', async () => {
    const host = embedding()
    const remove = vi.spyOn(window, 'removeEventListener')
    const pending = captureExcelContext(config, new AbortController().signal, 'unavailable')!
    expect(host.postMessage).toHaveBeenCalledWith(expect.objectContaining({ type: 'dsh/excel-context/request', version: 1 }), config.parentOrigin)
    const settled = vi.fn()
    void pending.then(settled)
    host.reply(host.envelope(), 'https://attacker.example')
    host.reply(host.envelope(), config.parentOrigin, window)
    host.reply({ ...host.envelope(), requestId: 'unrelated' })
    await Promise.resolve()
    expect(settled).not.toHaveBeenCalled()
    host.reply(host.envelope())
    expect(await pending).toBe(`[Excel context]\n${JSON.stringify(context)}\nCite Excel cells as [[cite:Sheet1!A1:B2]] (quote worksheet names with spaces).\n[/Excel context]`)
    expect(remove).toHaveBeenCalledWith('message', expect.any(Function))
  })

  it('pairs concurrent captures with their own frozen snapshots despite reversed responses', async () => {
    const host = embedding()
    const first = captureExcelContext(config, new AbortController().signal, 'unavailable')!
    const firstReply = host.envelope()
    const second = captureExcelContext(config, new AbortController().signal, 'unavailable')!
    host.reply(host.envelope({ ...context, selection: { ...context.selection, address: "'Forecast'!C5" } }))
    host.reply(firstReply)
    expect(await first).toContain("'Forecast'!A1:B4")
    expect(await second).toContain("'Forecast'!C5")
  })

  it('admits the exact UTF-8 JSON limit and refuses one additional byte', async () => {
    const host = embedding()
    const base = { ...context, workbook: { ...context.workbook, path: '' } }
    const bytes = new TextEncoder().encode(JSON.stringify(base)).byteLength
    const exact = { ...base, workbook: { ...base.workbook, path: 'x'.repeat(2048 - bytes) } }
    const accepted = captureExcelContext(config, new AbortController().signal, 'unavailable')!
    host.reply(host.envelope(exact))
    expect(await accepted).toContain(exact.workbook.path)
    const refused = captureExcelContext(config, new AbortController().signal, 'unavailable')!
    const failure = expect(refused).rejects.toThrow('unavailable')
    host.reply(host.envelope({ ...exact, workbook: { ...exact.workbook, path: `${exact.workbook.path}x` } }))
    await failure
  })

  it('rejects a cyclic structured-clone context without leaking a listener exception', async () => {
    const host = embedding()
    const cyclic: { self?: object } = {}
    cyclic.self = cyclic
    const pending = captureExcelContext(config, new AbortController().signal, 'unavailable')!
    const failure = expect(pending).rejects.toThrow('unavailable')
    host.reply(host.envelope(cyclic))
    await failure
  })

  it.each([
    { ...context, workbook: { ...context.workbook, path: '漢'.repeat(700) } },
    { ...context, selection: { ...context.selection, workbookId: 'other-book' } },
    { ...context, values: [[123]] },
    { ...context, observedAt: 'yesterday' },
    { ...context, worksheet: { name: 'Forecast' } },
  ])('fails closed for oversized or invalid context', async (payload) => {
    const host = embedding()
    const pending = captureExcelContext(config, new AbortController().signal, 'unavailable')!
    const failure = expect(pending).rejects.toThrow('unavailable')
    host.reply(host.envelope(payload))
    await failure
  })

  it('cleans up listeners on timeout, abort, and a parent failure', async () => {
    vi.useFakeTimers()
    const host = embedding()
    const remove = vi.spyOn(window, 'removeEventListener')
    const timeout = captureExcelContext(config, new AbortController().signal, 'unavailable')!
    const timedOut = expect(timeout).rejects.toThrow('unavailable')
    await vi.advanceTimersByTimeAsync(1000)
    await timedOut
    const controller = new AbortController()
    const pending = captureExcelContext(config, controller.signal, 'unavailable')!
    const cancelled = expect(pending).rejects.toThrow('unavailable')
    controller.abort()
    await cancelled
    const failed = captureExcelContext(config, new AbortController().signal, 'unavailable')!
    const rejected = expect(failed).rejects.toThrow('unavailable')
    const response = host.envelope()
    host.reply({ type: response.type, version: 1, requestId: response.requestId, error: 'context_unavailable' })
    await rejected
    expect(remove.mock.calls.filter(([name]) => name === 'message')).toHaveLength(3)
  })
})

describe('Excel cell citation binding', () => {
  it('decodes only the bounded canonical metadata at the end of its own message', async () => {
    const host = embedding()
    const pending = captureExcelContext(config, new AbortController().signal, 'unavailable')!
    host.reply(host.envelope())
    const marker = await pending
    expect(excelWorkbookOf(`First\n\n${marker}`)).toBe('book-1')
    expect(excelWorkbookOf('No captured metadata')).toBeUndefined()
    expect(excelWorkbookOf(`${marker}\nuser text`)).toBeUndefined()
    expect(excelWorkbookOf(marker.replace('book-1', 'wrong-id'))).toBeUndefined()
  })

  it('authenticates navigation acknowledgements and sends the historical workbook identity', async () => {
    const host = embedding()
    const pending = navigateExcelReference(config, 'historical-book', "'Forecast'!A1", new AbortController().signal)
    const request = host.postMessage.mock.calls[0]?.[0]
    if (request === undefined) throw new Error('missing navigation request')
    expect(request).toMatchObject({ type: 'dsh/excel-context/navigate', version: 1, workbookId: 'historical-book', address: "'Forecast'!A1" })
    const reply = { type: 'dsh/excel-context/navigation', version: 1, requestId: request.requestId, ok: true }
    const settled = vi.fn()
    void pending.then(settled)
    host.reply(reply, 'https://attacker.example')
    host.reply(reply, config.parentOrigin, window)
    await Promise.resolve()
    expect(settled).not.toHaveBeenCalled()
    host.reply(reply)
    await pending
    expect(settled).toHaveBeenCalledOnce()
  })
})
