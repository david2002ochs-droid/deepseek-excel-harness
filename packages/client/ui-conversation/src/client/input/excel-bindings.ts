import type { Context } from '@deepseek-ai/cordis'
import type { ConversationNodeDefinition } from '../contract/conversation.ts'
import type { UiConversation } from '../conversation/assembly.ts'
import { excelWorkbookOf } from './excel-context.ts'

interface InputBinding {
  readonly turn: number | undefined
  readonly workbookId: string | undefined
}

interface StepBinding extends InputBinding {
  readonly turn: number
  readonly step: number
}

const inputDefinition: ConversationNodeDefinition<InputBinding> = {
  kind: 'excel-input',
  match: event => event.type === 'user/message' && event.surfaceOp === 'append' && event.data.source.kind === 'user'
    ? { id: String(event.data.id), role: 'start' } : null,
  start: (_context, match) => {
    if (match.event.type !== 'user/message') throw new Error('excel-input requires user/message')
    const text = match.event.data.content.filter(block => block.type === 'text').map(block => block.text).join('\n')
    return { turn: match.location.kind === 'step' || match.location.kind === 'turn' ? match.location.turn.turn : undefined,
      workbookId: excelWorkbookOf(text) }
  },
  update: context => context.state,
}

const stepDefinition: ConversationNodeDefinition<StepBinding> = {
  kind: 'excelWorkbook',
  // The loop admits human steering before the next Step's request, so packed
  // Assistant history retains the same predecessor as the first live chunk.
  match: event => event.type === 'assistant/live-chunk' || event.type === 'assistant/attempt' || (event.type === 'assistant/message' && event.surfaceOp === 'append')
    ? { id: `${event.data.turn}:${event.data.step}`, role: 'start' } : null,
  start: (_context, match, reader) => {
    const event = match.event
    if (event.type !== 'assistant/live-chunk' && event.type !== 'assistant/attempt' && event.type !== 'assistant/message') throw new Error('excelWorkbook requires Assistant output')
    const input = reader.previous<InputBinding>('excel-input')?.state
    return { turn: event.data.turn, step: event.data.step,
      workbookId: input?.turn === event.data.turn ? input.workbookId : undefined }
  },
  update: context => context.state,
  buildLocationData: (context, scope, previous) => {
    if (scope !== 'step' || context.state === undefined) return null
    if (previous?.key === 'excelWorkbook' && previous.value === context.state) return previous
    return { kind: 'step', turn: context.state.turn, step: context.state.step, key: 'excelWorkbook', value: context.state }
  },
}

/**
 * Register immutable input-to-step binding projections for the plugin lifetime.
 * @param ctx - owning plugin context.
 * @param conversation - target-neutral Conversation registry.
 */
export function registerExcelBindings(ctx: Context, conversation: Pick<UiConversation, 'events'>): void {
  ctx.effect(() => conversation.events.register(inputDefinition))
  ctx.effect(() => conversation.events.register(stepDefinition))
}
