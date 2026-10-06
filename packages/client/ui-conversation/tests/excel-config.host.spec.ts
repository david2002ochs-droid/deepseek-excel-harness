import { Context } from '@deepseek-ai/cordis'
import type { IndexInjection } from '@deepseek-ai/dsh-host-webserver'
import { describe, expect, it } from 'vitest'
import * as HostPlugin from '../src/index.ts'
import { EXCEL_CONTEXT_BOOT_KEY, resolveExcelContextConfig } from '../src/excel-config.ts'

function collect(ctx: Context): IndexInjection[] {
  const rows: IndexInjection[] = []
  ctx.emit('webserver/index-inject', rows)
  return rows
}

describe('Excel deployment bootstrap', () => {
  it('supplies the Host opt-in before browser plugins start and removes it on unload', async () => {
    const ctx = new Context()
    const config = { parentOrigin: 'https://localhost:3443', timeoutMs: 10000 }
    const fiber = ctx.plugin(HostPlugin, { excelContext: config })
    await fiber.await()
    expect(collect(ctx)).toEqual([{ kind: 'global', name: EXCEL_CONTEXT_BOOT_KEY, value: config }])
    await fiber.dispose()
    expect(collect(ctx)).toEqual([])
  })

  it('leaves ordinary Web deployments without Excel boot state', async () => {
    const ctx = new Context()
    const fiber = ctx.plugin(HostPlugin)
    await fiber.await()
    expect(collect(ctx)).toEqual([])
    await fiber.dispose()
  })

  it('rejects invalid public boot values instead of silently disabling capture', () => {
    for (const parentOrigin of ['*', 'file:///workbook', 'https://localhost:3443/', 'https://user:pass@localhost:3443']) {
      expect(() => resolveExcelContextConfig({ parentOrigin, timeoutMs: 10000 })).toThrow()
    }
    expect(() => resolveExcelContextConfig({ parentOrigin: 'https://localhost:3443', timeoutMs: 0 })).toThrow()
    expect(() => resolveExcelContextConfig({ parentOrigin: 'https://localhost:3443' })).toThrow()
    expect(resolveExcelContextConfig(undefined)).toBeUndefined()
  })
})
