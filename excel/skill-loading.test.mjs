/** Actual catalog/tool/session replay with a scripted adapter; no network or paid model. */
import assert from 'node:assert/strict'
import { cp, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { test } from 'node:test'
import { Context } from '@deepseek-ai/cordis'
import { createUserMessage } from '@deepseek-ai/dsh-llm'
import { Session, SessionId } from '@deepseek-ai/dsh-session'
import SkillRegistry from '@deepseek-ai/dsh-skill'
import * as SkillFileSystem from '@deepseek-ai/dsh-skill-filesystem'
import * as ToolSkill from '@deepseek-ai/dsh-tool-skill'
import { mountAgentLoopTestDependencies, mountAgentLoopTestHarness } from '@deepseek-ai/dsh-agent-loop-testkit'
import { MockAdapter, textResponse, toolCallResponse } from '../packages/core/agent-loop/tests/mock-adapter.ts'

const directory = dirname(fileURLToPath(import.meta.url))
const references = ['code-analysis', 'debugging', 'forms', 'formulas', 'object-model-traps', 'pack', 'recovery', 'testing', 'xlflow-ui']

test('filesystem catalog and full skill tool results survive the production session replay', { timeout: 30000 }, async t => {
  const scratch = await mkdtemp(join(tmpdir(), 'excel-skill-loading-'))
  t.after(() => rm(scratch, { recursive: true, force: true }))
  const home = process.env.EXCEL_SKILL_TEST_HOME ?? join(scratch, 'home')
  if (!process.env.EXCEL_SKILL_TEST_HOME) {
    await mkdir(join(home, 'skills', 'xlflow', 'references'), { recursive: true })
    await cp(join(directory, 'skills', 'excel-vba'), join(home, 'skills', 'excel-vba'), { recursive: true })
    // Unit fixture stands in for the upstream bundle. Supply EXCEL_SKILL_TEST_HOME to check a real provisioned release.
    await writeFile(join(home, 'skills', 'xlflow', 'SKILL.md'), '---\nname: xlflow\ndescription: Scripted upstream fixture\n---\n\n# Fixture VBA workflow\nRead references/forms.md for forms.\n')
    for (const reference of references) await writeFile(join(home, 'skills', 'xlflow', 'references', `${reference}.md`), `# Fixture ${reference}\n`)
  }
  const ctx = new Context()
  t.after(() => ctx.fiber.dispose())
  await mountAgentLoopTestDependencies(ctx, { systemPrompt: { personaPrefix: '' } })
  await ctx.plugin(SkillRegistry)
  await ctx.plugin(SkillFileSystem, { dshHome: home, agentsHome: join(scratch, 'empty-agents-home'), watch: false })
  await ctx.plugin(ToolSkill)
  const driver = await mountAgentLoopTestHarness(ctx)
  const adapter = new MockAdapter([
    toolCallResponse('load-guard', 'skill', { name: 'excel-vba' }),
    toolCallResponse('load-upstream', 'skill', { name: 'xlflow' }),
    textResponse('Both skill instructions loaded.'),
  ])
  // The route label follows the owner; this adapter is synthetic and verifies no provider/network settings.
  ctx.llm.registerAdapter(['openrouter'], adapter)
  const agent = await driver.create(SessionId('excel-skill-loading'), {
    provider: 'openrouter', model: 'meta/muse-spark-1.3-contributor',
  }, { cwd: scratch })
  const idle = new Promise(resolve => {
    const dispose = ctx.on('agent/status', ({ agent: subject, status }) => {
      if (subject === agent && status === 'idle') { dispose(); resolve() }
    })
  })
  agent.followup(createUserMessage({ source: { kind: 'user' }, content: [{ type: 'text', text: 'Load Excel VBA guidance for this pane.' }] }))
  await idle
  const events = agent.session.snapshotEvents()
  const catalog = events.find(event => event.type === 'user/message' && event.data.source.kind === 'skill-catalog')
  assert.deepEqual(catalog.data.source.entries.map(entry => entry.name), ['excel-vba', 'xlflow'])
  const results = events.filter(event => event.type === 'tool/result')
  assert.equal(results.length, 2)
  assert.ok(results.every(event => event.data.message.isError === false))
  if (!process.env.EXCEL_SKILL_TEST_HOME) {
    const recorded = {
      catalog: catalog.data.source.entries,
      results: results.map((event, index) => ({
        role: event.data.message.role,
        source: event.data.message.source,
        toolCallId: event.data.message.toolCallId,
        isError: event.data.message.isError,
        content: event.data.message.content.map(block => ({
          ...block,
          // Only this fixture's resource prefix varies across machines; instruction text remains literal.
          text: block.text.replaceAll(join(home, 'skills', ['excel-vba', 'xlflow'][index]), `<home>/skills/${['excel-vba', 'xlflow'][index]}`),
        })),
      })),
    }
    const path = join(directory, 'fixtures', 'skill-loading.expected.json')
    if (process.env.RECORD_EXCEL_SKILL_SNAPSHOT === '1') await writeFile(path, `${JSON.stringify(recorded, null, 2)}\n`)
    assert.deepEqual(recorded, JSON.parse(await readFile(path, 'utf8')))
  }
  for (const [index, name] of ['excel-vba', 'xlflow'].entries()) {
    const definition = await ctx.skills.get(name, { cwd: scratch })
    const fullFile = await readFile(join(home, 'skills', name, 'SKILL.md'), 'utf8')
    const body = fullFile.replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n/u, '').trim()
    assert.equal(definition.content, body)
    assert.deepEqual(definition.resourceBase, { kind: 'directory', path: join(home, 'skills', name) })
    const content = results[index].data.message.content[0].text
    assert.ok(content.includes(body))
    assert.ok(content.includes(definition.resourceBase.path))
    assert.ok(JSON.stringify(adapter.requests[2].messages).includes(JSON.stringify(body).slice(1, -1)))
  }
  const upstream = await ctx.skills.get('xlflow', { cwd: scratch })
  for (const name of references) assert.ok((await readFile(join(upstream.resourceBase.path, 'references', `${name}.md`), 'utf8')).length > 0)
  const own = await ctx.skills.get('excel-vba', { cwd: scratch })
  const helper = await import(pathToFileURL(join(own.resourceBase.path, 'scripts', 'workbook.mjs')).href)
  await assert.rejects(helper.workbookRequest({ action: 'validate', context: {} }), { code: 'PANE_CONTEXT_REQUIRED' })
  const replay = Session.create(agent.id, JSON.parse(JSON.stringify(events)), agent.session.header)
  assert.deepEqual(replay.deriveMessages(), agent.session.deriveMessages())
  t.diagnostic(process.env.EXCEL_SKILL_TEST_HOME ? 'Verified real provisioned upstream release; scripted adapter, no model/network run.' : 'Verified own skill and upstream unit fixture; scripted adapter, no model/network run.')
})
