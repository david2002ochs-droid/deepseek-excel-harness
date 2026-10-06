/** Public deployment configuration supplied before the Conversation browser plugin starts. */
import z from '@deepseek-ai/schemastery'

/** Dedicated index-injection property; it carries no workbook data or credentials. */
export const EXCEL_CONTEXT_BOOT_KEY = '__DSH_EXCEL_CONTEXT__'

/** Excel embedding configuration; standalone documents do not capture context. */
export interface ExcelContextConfig {
  /** Exact trusted parent-frame origin. */
  readonly parentOrigin: string
  /** Maximum wait for the parent's metadata response, in milliseconds. */
  readonly timeoutMs: number
}

/** Host and browser use the same validation for this public configuration. */
export const ExcelContextConfigSchema: z<ExcelContextConfig> = z.object({
  parentOrigin: z.string().required(),
  timeoutMs: z.natural().min(1).required(),
})

/**
 * Decode the optional boot value and reject origins that contain a path or wildcard.
 * @param value - public index-injection value or explicit Client configuration.
 * @returns validated configuration, or undefined when the deployment did not opt in.
 */
export function resolveExcelContextConfig(value: unknown): ExcelContextConfig | undefined {
  if (value === undefined) return undefined
  if (typeof value !== 'object' || value === null
    || !('parentOrigin' in value) || typeof value.parentOrigin !== 'string'
    || !('timeoutMs' in value) || typeof value.timeoutMs !== 'number') {
    throw new Error('excelContext requires parentOrigin and timeoutMs')
  }
  const config = ExcelContextConfigSchema({ parentOrigin: value.parentOrigin, timeoutMs: value.timeoutMs })
  const origin = new URL(config.parentOrigin)
  if (origin.origin !== config.parentOrigin || !['https:', 'http:'].includes(origin.protocol)) {
    throw new Error('excelContext.parentOrigin must be an exact HTTP(S) origin')
  }
  return config
}

/**
 * Build the public boot injection shared by the Host and browser-loader fixture.
 * @param value - optional deployment configuration.
 * @returns the global injection, or undefined for ordinary web deployments.
 */
export function excelContextBootInjection(value: unknown):
  { kind: 'global'; name: typeof EXCEL_CONTEXT_BOOT_KEY; value: ExcelContextConfig } | undefined {
  const config = resolveExcelContextConfig(value)
  return config === undefined ? undefined : { kind: 'global', name: EXCEL_CONTEXT_BOOT_KEY, value: config }
}
