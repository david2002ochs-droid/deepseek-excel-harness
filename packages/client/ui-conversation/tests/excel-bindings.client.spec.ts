import { Context } from '@deepseek-ai/cordis'
import { Session, SessionId } from '@deepseek-ai/dsh-session'
import { createAssistantMessage, createUserMessage } from '@deepseek-ai/dsh-llm'
import { describe, expect, it } from 'vitest'
import { ConversationNodeAssembler } from '../src/client/conversation/assembler.ts'
import { ConversationEventRegistry } from '../src/client/conversation/event-registry.ts'
import type { ConversationTimelineSnapshot, ConversationViewDefinition } from '../src/client/contract/conversation.ts'
import { registerExcelBindings } from '../src/client/input/excel-bindings.ts'

const view: ConversationViewDefinition = {
  target: 'binding-test',
  create: () => ({ empty: { turnOrder: [], turns: new Map() },
    replace: input => input.timeline, apply: input => input.timeline }),
}

function marker(id: string): string {
  return `[Excel context]\n${JSON.stringify({ workbook: { id }, worksheet: { id: 'sheet-1', name: 'Sheet1' },
    selection: { workbookId: id, address: 'Sheet1!A1' }, observedAt: '2026-10-06T12:00:00.000Z' })}\nCite Excel cells as [[cite:Sheet1!A1:B2]] (quote worksheet names with spaces).\n[/Excel context]`
}

function fixture() {
  const ctx = new Context()
  const events = new ConversationEventRegistry(ctx)
  registerExcelBindings(ctx, { events })
  const assembler = new ConversationNodeAssembler(events, { entries: () => [view] })
  assembler.activateTarget('binding-test')
  const session = Session.create(SessionId('excel-bindings'))
  const entries = () => session.snapshotEvents().map(event => ({ type: 'event' as const, event }))
  const add = (workbookId?: string) => session.append('user/message', createUserMessage({ source: { kind: 'user' },
    content: [{ type: 'text', text: workbookId === undefined ? 'uncaptured' : `message\n\n${marker(workbookId)}` }] }), { surfaceOp: 'append' })
  const output = (turn: number, step: number) => session.append('assistant/message', {
    turn, step,
    message: createAssistantMessage({ content: [{ type: 'text', text: '[[cite:Sheet1!A1]]' }],
      source: { provider: 'openrouter', model: 'meta/muse-spark-1.3-contributor' } }),
    stream: [],
  }, { surfaceOp: 'append' })
  const binding = (turn: number, step: number) => {
    assembler.flush()
    const timeline = assembler.snapshot('binding-test') as ConversationTimelineSnapshot
    return timeline.turns.get(turn)?.steps.find(item => item.step === step)?.data.get('excelWorkbook')
  }
  return { ctx, assembler, session, entries, add, output, binding }
}

describe('incremental Excel response bindings', () => {
  it('freezes the first-output binding, adopts later steering only for the next step, and never borrows metadata', async () => {
    const f = fixture()
    try {
      f.session.append('turn/start', { turn: 1 }); f.add('book-a')
      f.session.append('step/start', { turn: 1, step: 1 }); f.output(1, 1)
      f.assembler.replaceWindow(f.entries(), false)
      const first = f.binding(1, 1)
      expect(first?.workbookId).toBe('book-a')
      f.add('book-b'); f.output(1, 1)
      for (const entry of f.entries().slice(4)) f.assembler.append(entry)
      expect(f.binding(1, 1)).toBe(first)
      f.session.append('step/end', { turn: 1, step: 1 })
      f.session.append('step/start', { turn: 1, step: 2 }); f.output(1, 2)
      for (const entry of f.entries().slice(6)) f.assembler.append(entry)
      expect(f.binding(1, 2)?.workbookId).toBe('book-b')
      f.add(); f.session.append('step/start', { turn: 1, step: 3 }); f.output(1, 3)
      for (const entry of f.entries().slice(9)) f.assembler.append(entry)
      expect(f.binding(1, 3)?.workbookId).toBeUndefined()
      f.session.append('turn/end', { turn: 1, reason: { kind: 'completed' } })
      f.session.append('turn/start', { turn: 2 }); f.session.append('step/start', { turn: 2, step: 1 }); f.output(2, 1)
      for (const entry of f.entries().slice(12)) f.assembler.append(entry)
      expect(f.binding(2, 1)?.workbookId).toBeUndefined()
      const expected = [f.binding(1, 1)?.workbookId, f.binding(1, 2)?.workbookId, f.binding(1, 3)?.workbookId, f.binding(2, 1)?.workbookId]
      f.assembler.replaceWindow(f.entries(), false)
      expect([f.binding(1, 1)?.workbookId, f.binding(1, 2)?.workbookId,
        f.binding(1, 3)?.workbookId, f.binding(2, 1)?.workbookId]).toEqual(expected)
    } finally { await f.ctx.fiber.dispose() }
  })

  it('republishes missing binding when older input is prepended and matches a complete replay', async () => {
    const f = fixture()
    try {
      f.session.append('turn/start', { turn: 1 }); f.add('historical-book')
      f.session.append('step/start', { turn: 1, step: 1 }); f.output(1, 1)
      const entries = f.entries()
      f.assembler.replaceWindow(entries.slice(2), true)
      expect(f.binding(1, 1)?.workbookId).toBeUndefined()
      f.assembler.prepend(entries.slice(0, 2), false)
      expect(f.binding(1, 1)?.workbookId).toBe('historical-book')
      f.assembler.replaceWindow(entries, false)
      expect(f.binding(1, 1)?.workbookId).toBe('historical-book')
    } finally { await f.ctx.fiber.dispose() }
  })
})
