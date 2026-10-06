// @vitest-environment jsdom
import { readFileSync } from 'node:fs'
import { load } from 'js-yaml'
import { expect, it, vi } from 'vitest'
import { RemoteMock } from '@deepseek-ai/dsh-remote-mock'
import { TestClient, remoteDefaultResponses, webApp } from '@deepseek-ai/dsh-client-test-runtime/src/assembly/index.ts'
import { Session, SessionId } from '@deepseek-ai/dsh-session'
import { createUserMessage } from '@deepseek-ai/dsh-llm'
import { SessionInputShell } from '../src/client/input/facade.ts'
import type { ConversationConfig } from '../src/client/index.ts'
import { EXCEL_CONTEXT_BOOT_KEY, excelContextBootInjection } from '../src/excel-config.ts'

it('receives the Host boot opt-in without per-entry Client config and retains captured context in model history', async () => {
  const rows = load(readFileSync('packages/client/ui-conversation/tests/fixtures/excel-context.cordis.yml', 'utf8')) as
    { name: string; config: ConversationConfig }[]
  const row = rows[0]!
  const boot = excelContextBootInjection(row.config.excelContext)
  if (boot === undefined) throw new Error('Host did not supply the Excel opt-in')
  const previous = Object.getOwnPropertyDescriptor(globalThis, EXCEL_CONTEXT_BOOT_KEY)
  Object.defineProperty(globalThis, EXCEL_CONTEXT_BOOT_KEY, { configurable: true, value: boot.value })
  try {
    const client = await TestClient.start({ roster: webApp.closure([row.name]) }, RemoteMock.create().load(remoteDefaultResponses))
    const parentFrame = document.createElement('iframe')
    document.body.appendChild(parentFrame)
    const parent = parentFrame.contentWindow!
    const postMessage = vi.spyOn(parent, 'postMessage').mockImplementation(() => {})
    const spy = vi.spyOn(window, 'parent', 'get').mockReturnValue(parent)
    let shell: SessionInputShell | undefined
    try {
      const entry = [...client.ctx.loader.entries()].find(candidate => candidate.options.name === row.name)!
      expect(entry.fiber).toBeDefined()
      // Production ClientEntries creates package rows without copying Host config.
      expect(entry.options.config).toBeUndefined()
      const session = Session.create(SessionId('loader-excel-context'))
      shell = new SessionInputShell({ actx: client.ctx,
        commandAttachments: { serialize: async () => [], release: () => {}, unsupportedNotice: () => '' },
        captureMessageContext: signal => client.ctx.bail(client.ctx, 'conversation/message-context', signal),
        defaultSink: async (text) => {
          session.append('user/message', createUserMessage({ source: { kind: 'user' }, content: [{ type: 'text', text }] }), { surfaceOp: 'append' })
          return { kind: 'success' }
        },
      })
      shell.setDraft('Summarize selection'); shell.submit()
      expect(postMessage).toHaveBeenCalledOnce()
      const request: unknown = postMessage.mock.calls[0]![0]
      if (typeof request !== 'object' || request === null || !('requestId' in request) || typeof request.requestId !== 'string') {
        throw new Error('Missing Excel capture request identity')
      }
      const requestId = request.requestId
      const context = { workbook: { id: 'book-1', name: 'Budget 2027.xlsx', path: 'C:\\Models\\Budget 2027.xlsx' }, worksheet: { id: 'sheet-1', name: 'Sheet1' },
        selection: { workbookId: 'book-1', address: 'Sheet1!A1' }, observedAt: '2026-10-06T12:00:00.000Z' }
      window.dispatchEvent(new MessageEvent('message', { source: parent, origin: row.config.excelContext!.parentOrigin,
        data: { type: 'dsh/excel-context/response', version: 1, requestId, context } }))
      await vi.waitFor(() => { expect(session.deriveMessages()).toHaveLength(1) })
      const replay = Session.create(SessionId('loader-excel-replayed'), session.snapshotEvents())
      expect(replay.deriveMessages()[0]!.content).toEqual([{ type: 'text', text:
        `Summarize selection\n\n[Excel context]\n${JSON.stringify(context)}\nCite Excel cells as [[cite:Sheet1!A1:B2]] (quote worksheet names with spaces).\n[/Excel context]` }])
      await client.unload(row.name)
      expect(client.ctx.bail(client.ctx, 'conversation/message-context', new AbortController().signal)).toBeUndefined()
    } finally {
      shell?.dispose(); spy.mockRestore(); postMessage.mockRestore(); parentFrame.remove(); await client.dispose()
    }
  } finally {
    if (previous === undefined) Reflect.deleteProperty(globalThis, EXCEL_CONTEXT_BOOT_KEY)
    else Object.defineProperty(globalThis, EXCEL_CONTEXT_BOOT_KEY, previous)
  }
}, 60_000)
