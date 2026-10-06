/** Keyless private-home integration, preservation and cancellation checks. */
import assert from 'node:assert/strict'
import { execFile } from 'node:child_process'
import { lstat, mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { delimiter, join } from 'node:path'
import { setTimeout as delay } from 'node:timers/promises'
import { test } from 'node:test'
import { fileURLToPath } from 'node:url'
import { promisify } from 'node:util'
import { runtimeEnvironment, setupExcelHome, validateExcelHome } from './setup.mjs'

async function fixture(t) {
  const home = await mkdtemp(join(tmpdir(), 'excel-setup-test-'))
  t.after(() => rm(home, { recursive: true, force: true }))
  const provision = async (target, { signal }) => {
    signal.throwIfAborted()
    await mkdir(join(target, 'runtime', 'xlflow-0.35.0'), { recursive: true })
    await mkdir(join(target, 'skills', 'xlflow'), { recursive: true })
    return { directory: join(target, 'runtime', 'xlflow-0.35.0') }
  }
  return { home, provision, setup: options => setupExcelHome(home, { provision, ...options }) }
}

test('installs relocatable resources once and preserves skills, settings and existing instructions', async t => {
  const f = await fixture(t)
  await mkdir(join(f.home, 'skills', 'user-skill'), { recursive: true })
  await writeFile(join(f.home, 'skills', 'user-skill', 'SKILL.md'), 'user data')
  await writeFile(join(f.home, 'AGENTS.md'), 'owner instructions\r\n')
  await writeFile(join(f.home, 'settings.json'), '{"owner":"settings"}')
  const result = await f.setup()
  const original = await readFile(join(f.home, 'AGENTS.md'), 'utf8')
  assert.ok(original.startsWith('owner instructions\r\n'))
  assert.match(original, /load the excel-vba skill before xlflow/u)
  await writeFile(join(result.excelVbaSkillRoot, 'SKILL.md'), 'owner customization')
  await f.setup()
  assert.equal(await readFile(join(result.excelVbaSkillRoot, 'SKILL.md'), 'utf8'), 'owner customization')
  assert.equal(await readFile(join(f.home, 'AGENTS.md'), 'utf8'), original)
  assert.equal(await readFile(join(f.home, 'settings.json'), 'utf8'), '{"owner":"settings"}')
  assert.equal(await readFile(join(f.home, 'skills', 'user-skill', 'SKILL.md'), 'utf8'), 'user data')
  assert.ok((await readFile(join(result.excelVbaSkillRoot, 'scripts', 'workbook.mjs'), 'utf8')).includes('workbookRequest'))
  assert.ok(!(await readdir(f.home)).includes('.excel-setup.lock'))
})

test('unmanaged skills and edited managed instructions fail without replacement', async t => {
  const f = await fixture(t)
  await mkdir(join(f.home, 'skills', 'excel-vba'), { recursive: true })
  await writeFile(join(f.home, 'skills', 'excel-vba', 'SKILL.md'), 'unmanaged')
  await assert.rejects(f.setup(), { code: 'SKILL_COLLISION' })
  assert.equal(await readFile(join(f.home, 'skills', 'excel-vba', 'SKILL.md'), 'utf8'), 'unmanaged')
  const g = await fixture(t)
  await g.setup()
  const path = join(g.home, 'AGENTS.md')
  const edited = (await readFile(path, 'utf8')).replace('before xlflow', 'custom instruction')
  await writeFile(path, edited)
  await assert.rejects(g.setup(), { code: 'INSTRUCTION_CHANGED' })
  assert.equal(await readFile(path, 'utf8'), edited)
})

test('same-home provisioning is serialized and the lock is released on failure', async t => {
  const f = await fixture(t)
  let active = 0
  let maximum = 0
  const provision = async (...args) => {
    active++
    maximum = Math.max(maximum, active)
    await delay(80)
    active--
    return f.provision(...args)
  }
  await Promise.all([f.setup({ provision }), f.setup({ provision })])
  assert.equal(maximum, 1)
  await assert.rejects(f.setup({ provision: async () => { throw new Error('fixture failure') } }), /fixture failure/u)
  assert.ok(!(await readdir(f.home)).includes('.excel-setup.lock'))
  await f.setup()
})

test('abort waits for provisioning cleanup and starts no skill/instruction step', async t => {
  const f = await fixture(t)
  const controller = new AbortController()
  let stopped = false
  let started
  const entered = new Promise(resolve => { started = resolve })
  const provision = async (_home, { signal }) => {
    started()
    try { await delay(10000, undefined, { signal }) }
    finally { await delay(20); stopped = true }
  }
  const pending = f.setup({ provision, signal: controller.signal })
  await entered
  controller.abort()
  await assert.rejects(pending, { name: 'AbortError' })
  assert.equal(stopped, true)
  assert.deepEqual(await readdir(f.home), [])
})

test('cancellation during home validation prevents setup writes and provisioning', async t => {
  const f = await fixture(t)
  const controller = new AbortController()
  const home = join(f.home, 'missing', 'home')
  let calls = 0
  const pending = setupExcelHome(home, { signal: controller.signal, provision: async () => { calls++ } })
  controller.abort()
  await assert.rejects(pending, { name: 'AbortError' })
  assert.equal(calls, 0)
  assert.deepEqual(await readdir(f.home), [])
})

test('unavailable drive roots reject setup and actual launcher before writes or service startup', { skip: process.platform !== 'win32' }, async t => {
  let drive
  for (const letter of 'ZYXWVUTSRQPONMLKJIHGFEDCBA') {
    const root = `${letter}:\\`
    try { await lstat(root) }
    catch (error) { if (error.code !== 'ENOENT') throw error; drive = root; break }
  }
  if (!drive) { t.skip('No unavailable drive root on this host.'); return }
  const home = join(drive, 'excel-setup-unavailable-root')
  await assert.rejects(lstat(home), { code: 'ENOENT' })
  await assert.rejects(validateExcelHome(home), { code: 'HOME_UNAVAILABLE' })
  let calls = 0
  await assert.rejects(setupExcelHome(home, { provision: async () => { calls++ } }), { code: 'HOME_UNAVAILABLE' })
  assert.equal(calls, 0)
  const launcher = fileURLToPath(new URL('./start.mjs', import.meta.url))
  await assert.rejects(promisify(execFile)(process.execPath, [launcher], {
    env: { ...process.env, EXCEL_DSH_HOME: home, npm_execpath: 'unused-npm-entry' },
    windowsHide: true, timeout: 3000,
  }), error => {
    assert.equal(error.code, 1)
    assert.equal(error.killed, false)
    assert.match(error.stderr, /Excel entry setup failed/u)
    return true
  })
  await assert.rejects(lstat(home), { code: 'ENOENT' })
})

test('lock waiting is bounded and cancelled waiters do not remove another owner lock', async t => {
  const f = await fixture(t)
  await mkdir(join(f.home, '.excel-setup.lock'))
  let calls = 0
  await assert.rejects(f.setup({ timeoutMs: 80, provision: async () => { calls++ } }), { name: 'AbortError' })
  assert.equal(calls, 0)
  assert.deepEqual(await readdir(f.home), ['.excel-setup.lock'])
  const controller = new AbortController()
  controller.abort()
  await assert.rejects(f.setup({ signal: controller.signal }), { name: 'AbortError' })
})

test('private-home validation and runtime environment preserve all other values', async t => {
  await assert.rejects(validateExcelHome('relative-home'), { code: 'HOME_REQUIRED' })
  await assert.rejects(validateExcelHome(process.cwd()), { code: 'HOME_IN_CHECKOUT' })
  const f = await fixture(t)
  const env = { Path: 'existing-path', MODEL_ROUTE: 'owner-route', DSH_HOME: f.home, SECURITY: 'confined' }
  const runtime = await f.setup()
  assert.deepEqual(runtimeEnvironment(env, runtime), { ...env, Path: `${runtime.directory}${delimiter}existing-path` })
  assert.equal(env.Path, 'existing-path')
})
