/** Captured Excel metadata survives Web submission, model replay, persistence, and citation navigation. */
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { createServer, type Server } from 'node:http'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import yaml from 'js-yaml'
import { composeEntries, loadOverlayPatches } from '@deepseek-ai/dsh-app-boot'
import { entryListSchema } from '@deepseek-ai/cordis-plugin-include'
import { chromium, type Browser, type Page } from 'playwright'
import { afterEach, describe, expect, it, onTestFailed } from 'vitest'
import {
  assertFixtureInventory, compareOrRefreshGolden, fixtureUserPrompts,
  launchWebScaffold, webSnapshotMode, type WebScaffold,
} from './scaffold.ts'
import { connectFreshWorkspace, newEnglishPage, saveFailureShot } from './support.ts'

const DIRECTORY = fileURLToPath(new URL('../../../snapshots/web/excel-context', import.meta.url))
const FIXTURE = join(DIRECTORY, 'session.v4.jsonl')
const OVERLAY = fileURLToPath(new URL('./excel-context.overlay.yml', import.meta.url))
const PRESET = fileURLToPath(new URL('../../../packages/bundle/web-app/presets/standard.patch.yml', import.meta.url))
const MODE = webSnapshotMode()

describe.skipIf(MODE === 'record')('web e2e: recorded Excel context', () => {
  let scaffold: WebScaffold | undefined
  let browser: Browser | undefined
  let page: Page
  let parent: Server | undefined
  let directory: string | undefined

  afterEach(async () => {
    try { await browser?.close() } finally {
      try { await scaffold?.close() } finally {
        try {
          if (parent?.listening === true) await new Promise<void>((resolve, reject) => {
            parent!.close((error) => { if (error === undefined) resolve(); else reject(error) })
          })
        } finally {
          if (directory !== undefined) await rm(directory, { recursive: true, force: true })
        }
      }
    }
  })

  it('captures the trusted parent and reopens the recorded cell in its originating workbook', async () => {
    const recorded = fixtureUserPrompts(await readFile(FIXTURE, 'utf8'))[0]!
    const marker = recorded.indexOf('[Excel context]\n')
    const prompt = recorded.slice(0, marker).trimEnd()
    const context = JSON.parse(recorded.slice(marker).split('\n')[1]!) as Record<string, unknown>
    parent = createServer((_request, response) => {
      response.writeHead(200, { 'content-type': 'text/html' })
      response.end('<!doctype html><html><body><iframe title="Harness" style="width:1600px;height:950px;border:0"></iframe></body></html>')
    })
    await new Promise<void>((resolve, reject) => { parent!.once('error', reject); parent!.listen(0, '127.0.0.1', resolve) })
    const address = parent.address()
    if (address === null || typeof address === 'string') throw new Error('Office fixture requires a TCP port')
    const parentOrigin = `http://127.0.0.1:${address.port}`
    directory = await mkdtemp(join(tmpdir(), 'dsh-excel-context-'))
    const overlay = join(directory, 'overlay.yml')
    const preset = composeEntries([loadOverlayPatches('Excel snapshot', PRESET)])[0]!
    const config = preset.config as { id: string; order: number; plugins: { id: string }[] }
    // This text-only fixture excludes OS-dependent shells and the Windows ACL skill catalog.
    const plugins = config.plugins.filter(plugin => !['tool-skill', 'tool-bash', 'tool-pwsh'].includes(plugin.id))
    const presetPatch = yaml.dump([{ id: preset.id, config: { ...config, plugins } }], { schema: entryListSchema })
    await writeFile(overlay, (await readFile(OVERLAY, 'utf8')).replace('http://127.0.0.1:3443', parentOrigin) + presetPatch)
    scaffold = await launchWebScaffold({
      replayFixture: FIXTURE, extraOverlayPath: overlay,
      replayProviders: [{ id: 'meta', name: 'Meta', models: [
        { id: 'meta/muse-spark-1.3-contributor', name: 'Muse Spark', contextWindow: 128000, defaultMaxTokens: 8192 },
      ] }],
    })
    const requests: string[] = []
    scaffold.ctx.on('llm/stream', async function* (options, next) {
      expect({ provider: options.provider, model: options.model, maxTokens: options.maxTokens }).toEqual({
        provider: 'meta', model: 'meta/muse-spark-1.3-contributor', maxTokens: 8192,
      })
      requests.push(JSON.stringify(options.messages))
      yield* next()
    })
    browser = await chromium.launch()
    page = await newEnglishPage(browser)
    const errors: string[] = []
    page.on('requestfailed', request => errors.push(request.failure()?.errorText ?? 'Request failed'))
    page.on('pageerror', error => errors.push(error.message.replace(/https?:\/\/\S+/gu, '[URL]')))
    page.on('console', (message) => { if (message.type() === 'error') errors.push(message.text().replace(/https?:\/\/\S+/gu, '[URL]')) })
    onTestFailed(() => saveFailureShot(page, 'web-excel-context'))
    await page.goto(`${parentOrigin}/office`)
    await page.evaluate(({ url, metadata }) => {
      const frame = document.querySelector('iframe')!
      window.addEventListener('message', (event) => {
        if (event.source !== frame.contentWindow || event.origin !== new URL(url).origin) return
        const request = event.data as { type: string; requestId: string; workbookId?: string; address?: string }
        if (request.type === 'dsh/excel-context/request') {
          frame.contentWindow!.postMessage({ type: 'dsh/excel-context/response', version: 1,
            requestId: request.requestId, context: metadata }, event.origin)
        } else if (request.type === 'dsh/excel-context/navigate') {
          document.body.dataset.navigation = JSON.stringify({ workbookId: request.workbookId, address: request.address })
          frame.contentWindow!.postMessage({ type: 'dsh/excel-context/navigation', version: 1,
            requestId: request.requestId, ok: true }, event.origin)
        }
      })
      frame.src = url
    }, { url: scaffold.authenticatedUrl, metadata: context })
    const app = page.frameLocator('iframe')
    try { await connectFreshWorkspace(app, scaffold.workspaceCwd) } catch (error) {
      await saveFailureShot(page, 'web-excel-context')
      expect(errors).toEqual([])
      throw error
    }
    const input = app.locator('[data-composer-input]').first()
    await input.fill(prompt)
    await Promise.all([scaffold.whenTurnSettled(), input.press('Enter')])
    expect(requests).toHaveLength(1)
    expect(requests[0]).toContain(JSON.stringify(recorded).slice(1, -1))
    const citation = app.getByRole('button', { name: 'Sheet1!A1:B2', exact: true })
    await citation.waitFor()
    await compareOrRefreshGolden(join(DIRECTORY, 'citation.expected.md'),
      await citation.locator('..').ariaSnapshot(), MODE)
    await citation.click()
    await expect.poll(() => page.locator('body').getAttribute('data-navigation')).toBe(
      JSON.stringify({ workbookId: 'recorded-book', address: 'Sheet1!A1:B2' }))
    expect(errors).toEqual([])
    await assertFixtureInventory(DIRECTORY,
      ['session.v4.jsonl', 'citation.expected.md', 'system-prompt.expected.md', 'tool-schemas.expected.json'])
  }, 120000)
})
