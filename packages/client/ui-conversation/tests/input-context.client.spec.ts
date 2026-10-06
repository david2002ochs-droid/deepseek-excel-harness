import { describe, expect, it, vi } from 'vitest'
import type { Context } from '@deepseek-ai/cordis'
import type { InputTriggerController, PickOutcome, SubmitOutcome } from '../src/client/contract/input.ts'
import { SessionInputShell } from '../src/client/input/facade.ts'
import { Session, SessionId } from '@deepseek-ai/dsh-session'
import { createUserMessage } from '@deepseek-ai/dsh-llm'

const commandAttachments = { serialize: async () => [], release: () => {}, unsupportedNotice: () => '' }
const triggers = (adjudicate: InputTriggerController['adjudicate']): InputTriggerController => ({
  adjudicate, launcher: { getSnapshot: () => null, subscribe: () => () => {} },
  lexicon: { getSnapshot: () => new Map(), subscribe: () => () => {} },
  track: () => {}, arbitrate: () => 'pass', onSpace: () => false,
  serializeReference: async () => '', openReference: () => false, toggleSource: () => {},
})

describe('composer message context', () => {
  it('captures before slash arbitration and retains the same context until the message sends', async () => {
    const decision = Promise.withResolvers<PickOutcome>()
    const capture = Promise.withResolvers<string>()
    const captured = vi.fn(() => capture.promise)
    const sink = vi.fn(async (_text: string): Promise<SubmitOutcome> => ({ kind: 'success' }))
    const shell = new SessionInputShell({ actx: {} as Context, commandAttachments, defaultSink: sink,
      captureMessageContext: captured, inputTriggers: () => triggers(() => decision.promise) })
    try {
      shell.setDraft('/ordinary')
      shell.submit()
      expect(captured).toHaveBeenCalledOnce()
      decision.resolve(undefined)
      await Promise.resolve()
      expect(sink).not.toHaveBeenCalled()
      capture.resolve('selection at gesture')
      await vi.waitFor(() => { expect(sink).toHaveBeenCalledWith('/ordinary\n\nselection at gesture', [], 'queue', expect.any(AbortSignal)) })
      expect(captured).toHaveBeenCalledOnce()
    } finally { shell.dispose() }
  })

  it('pairs each detached message with its captured context and preserves ordinary messages', async () => {
    const first = Promise.withResolvers<string>()
    const second = Promise.withResolvers<string>()
    const capture = vi.fn().mockReturnValueOnce(first.promise).mockReturnValueOnce(second.promise).mockReturnValueOnce(undefined)
    const sink = vi.fn(async (_text: string): Promise<SubmitOutcome> => ({ kind: 'success' }))
    const shell = new SessionInputShell({ actx: {} as Context, commandAttachments, defaultSink: sink, captureMessageContext: capture })
    try {
      shell.setDraft('first'); shell.submit()
      shell.setDraft('second'); shell.submit('steer')
      second.resolve('C5')
      first.resolve('A1:B4')
      await vi.waitFor(() => { expect(sink).toHaveBeenCalledTimes(2) })
      expect(sink.mock.calls.map(([text]) => text).sort()).toEqual(['first\n\nA1:B4', 'second\n\nC5'])
      shell.setDraft('standalone'); shell.submit()
      expect(sink).toHaveBeenLastCalledWith('standalone', [], 'queue', expect.any(AbortSignal))
    } finally { shell.dispose() }
  })

  it('restores a failed capture and never admits metadata from a disposed submission', async () => {
    const capture = Promise.withResolvers<string>()
    const sink = vi.fn()
    const captured = vi.fn<(signal: AbortSignal) => Promise<string>>()
      .mockRejectedValueOnce(new Error('capture unavailable')).mockReturnValueOnce(capture.promise)
    const shell = new SessionInputShell({ actx: {} as Context, commandAttachments, defaultSink: sink, captureMessageContext: captured })
    shell.setDraft('preserve me'); shell.submit()
    await vi.waitFor(() => { expect(shell.snapshot.draft).toBe('preserve me') })
    expect(shell.notices.getSnapshot()).toMatchObject({ level: 'error', text: 'capture unavailable' })
    shell.submit()
    const signal = captured.mock.calls[1]?.[0] as AbortSignal
    shell.dispose()
    expect(signal.aborted).toBe(true)
    capture.resolve('late')
    await Promise.resolve()
    expect(sink).not.toHaveBeenCalled()
  })

  it('reconstructs captured workbook/selection context from durable user-message history', async () => {
    const context = '[Excel context]\n{"workbook":{"id":"book-1"},"worksheet":{"id":"sheet-1","name":"Sheet1"},"selection":{"workbookId":"book-1","address":"Sheet1!A1"},"observedAt":"2026-10-06T12:00:00.000Z"}\nCite Excel cells as [[cite:Sheet1!A1:B2]] (quote worksheet names with spaces).\n[/Excel context]'
    const session = Session.create(SessionId('excel-context'))
    const shell = new SessionInputShell({ actx: {} as Context, commandAttachments,
      captureMessageContext: async () => context,
      defaultSink: async (text) => {
        session.append('user/message', createUserMessage({ source: { kind: 'user' }, content: [{ type: 'text', text }] }), { surfaceOp: 'append' })
        return { kind: 'success' }
      },
    })
    try {
      shell.setDraft('Summarize selection'); shell.submit()
      await vi.waitFor(() => { expect(session.deriveMessages()).toHaveLength(1) })
      const restored = Session.create(SessionId('excel-restored'), session.snapshotEvents())
      expect(restored.deriveMessages()[0]?.content).toEqual([{ type: 'text', text: `Summarize selection\n\n${context}` }])
      expect(restored.deriveMessages()[0]?.content).toMatchInlineSnapshot(`
        [
          {
            "text": "Summarize selection

        [Excel context]
        {\"workbook\":{\"id\":\"book-1\"},\"worksheet\":{\"id\":\"sheet-1\",\"name\":\"Sheet1\"},\"selection\":{\"workbookId\":\"book-1\",\"address\":\"Sheet1!A1\"},\"observedAt\":\"2026-10-06T12:00:00.000Z\"}
        Cite Excel cells as [[cite:Sheet1!A1:B2]] (quote worksheet names with spaces).
        [/Excel context]",
            "type": "text",
          },
        ]
      `)
    } finally { shell.dispose() }
  })
})
