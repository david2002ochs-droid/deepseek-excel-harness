/** Host registration for browser conversation preferences. */
import type {} from '@deepseek-ai/dsh-settings'
import type {} from '@deepseek-ai/dsh-host-webserver'

import type { Volatile, Context } from '@deepseek-ai/cordis'
import type { BusyEnterBehavior } from './submission-settings.ts'
import z from '@deepseek-ai/schemastery'
import { BUSY_ENTER_FIELD } from './submission-settings.ts'

import { ConversationSettingsFields } from './submission-settings.ts'
import {
  ExcelContextConfigSchema, excelContextBootInjection, type ExcelContextConfig,
} from './excel-config.ts'

export {
  BUSY_ENTER_BEHAVIORS, BUSY_ENTER_FIELD, CONVERSATION_SETTINGS_NAMESPACE,
  DEFAULT_BUSY_ENTER_BEHAVIOR, type BusyEnterBehavior, type ConversationSettings,
} from './submission-settings.ts'

/** Runtime preferences projected to the browser. */
export interface Config {
  /** Enter key behavior while a turn is running. */
  busyEnter: Volatile<BusyEnterBehavior>
  /** Public opt-in supplied to the browser before its plugins start. */
  excelContext?: ExcelContextConfig | undefined
}

/** Live preferences projected to the browser. */
export const Config = z.object({
  busyEnter: ConversationSettingsFields[BUSY_ENTER_FIELD].volatile(),
  excelContext: z.union([z.const(undefined), ExcelContextConfigSchema]),
})

/** Host preferences are consumed through the configuration form projection.
 * @param ctx Plugin context used for optional settings presentation.
 * @param config Validated Host preferences and Excel deployment configuration.
 */
export function apply(ctx: Context, config: Config): void {
  const boot = excelContextBootInjection(config.excelContext)
  ctx.inject(['settings'], (child) => { child.effect(() => child.settings.configure({ auto: false }, ctx.fiber)) })
  if (boot !== undefined) {
    ctx.on('webserver/index-inject', (table) => {
      table.push(boot)
    })
  }
}
