/** Keyless provisioning tests; fake release bytes never launch native executables. */
import assert from 'node:assert/strict'
import { cp, mkdtemp, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { test } from 'node:test'
import { provisionXlflow } from './xlflow-runtime.mjs'

const references = ['code-analysis', 'debugging', 'forms', 'formulas', 'object-model-traps',
  'pack', 'recovery', 'testing', 'xlflow-ui']

test('abort during download stops every subsequent operation and removes owned staging', async t => {
  const controller = new AbortController()
  let extraction = 0
  const f = await fixture(t, {
    signal: controller.signal,
    async download(_path, signal) { assert.equal(signal, controller.signal); controller.abort() },
    async extract() { extraction++ },
  })
  await assert.rejects(f.setup(), { name: 'AbortError' })
  assert.equal(extraction, 0)
  assert.deepEqual(await readdir(join(f.home, 'runtime')), [])
})

test('native version-child cancellation awaits close before removing its executable staging', { skip: process.platform !== 'win32' }, async t => {
  const controller = new AbortController()
  const f = await fixture(t, {
    signal: controller.signal,
    async extract(_archive, directory) {
      // A copied Windows command interpreter waits on stdin; no VBA or upstream CLI is invoked.
      await cp(join(process.env.SystemRoot, 'System32', 'cmd.exe'), join(directory, 'xlflow.exe'))
      await writeFile(join(directory, 'xlflow-excel-bridge.exe'), 'fixture')
      await writeFile(join(directory, 'LICENSE'), 'fixture')
    },
    async verifyFiles() {
      const timer = setTimeout(() => controller.abort(), 100)
      t.after(() => clearTimeout(timer))
    },
  })
  delete f.operations.run
  await assert.rejects(f.setup(), { code: 'PROCESS_FAILED' })
  assert.deepEqual(await readdir(join(f.home, 'runtime')), [])
})

async function fixture(t, overrides = {}) {
  const home = await mkdtemp(join(tmpdir(), 'xlflow-runtime-test-'))
  t.after(() => rm(home, { recursive: true, force: true }))
  const calls = []
  const operations = {
    async download(path) { calls.push('download'); await writeFile(path, 'fake archive') },
    async verifyArchive() {},
    async extract(archive, directory) {
      calls.push('extract')
      for (const name of ['xlflow.exe', 'xlflow-excel-bridge.exe', 'LICENSE']) {
        await writeFile(join(directory, name), 'fixture')
      }
    },
    async verifyFiles(directory) {
      for (const name of ['xlflow.exe', 'xlflow-excel-bridge.exe', 'LICENSE']) {
        await readFile(join(directory, name))
      }
    },
    async run(command, args) {
      calls.push(args.join(' '))
      if (args[0] === 'version') return JSON.stringify({ status: 'ok', version: { version: '0.35.0' } })
      if (basename(command) === 'xlflow-excel-bridge.exe') {
        return JSON.stringify({ name: 'xlflow-excel-bridge', version: '1.0.0', protocol_version: 1 })
      }
      const skill = join(args[args.indexOf('--target') + 1], 'xlflow')
      await mkdir(join(skill, 'references'), { recursive: true })
      await writeFile(join(skill, 'SKILL.md'), 'upstream skill')
      for (const name of references) await writeFile(join(skill, 'references', `${name}.md`), `reference ${name}`)
      return JSON.stringify({ status: 'ok' })
    },
    ...overrides,
  }
  return { home, calls, operations, setup: () => provisionXlflow(home, operations) }
}

test('checksum mismatch prevents extraction and execution and removes staging', async t => {
  const f = await fixture(t, { verifyArchive: undefined })
  delete f.operations.verifyArchive
  await assert.rejects(f.setup(), { code: 'CHECKSUM_MISMATCH' })
  assert.deepEqual(f.calls, ['download'])
  assert.deepEqual(await readdir(join(f.home, 'runtime')), [])
})

test('missing bridge prevents executing the CLI', async t => {
  const f = await fixture(t, {
    async extract(archive, directory) { await writeFile(join(directory, 'xlflow.exe'), 'fixture') },
    verifyFiles: undefined,
  })
  delete f.operations.verifyFiles
  await assert.rejects(f.setup(), error => error.code === 'RUNTIME_INVALID' && error.message.includes('xlflow-excel-bridge.exe'))
  assert.deepEqual(f.calls, ['download'])
  assert.deepEqual(await readdir(join(f.home, 'runtime')), [])
})

test('CLI and bridge version mismatches are rejected before publishing the runtime', async t => {
  for (const target of ['version', '--version-json']) {
    const f = await fixture(t)
    const original = f.operations.run
    f.operations.run = async (command, args) => args[0] === target ? '{"status":"ok","version":{"version":"0.34.0"}}' : original(command, args)
    await assert.rejects(f.setup(), { code: 'VERSION_MISMATCH' })
    assert.deepEqual(await readdir(join(f.home, 'runtime')), [])
  }
})

test('successful setup installs all references in the actual Harness skill root and retains license', async t => {
  const f = await fixture(t)
  const result = await f.setup()
  assert.equal(result.version, '0.35.0')
  assert.equal(result.skillRoot, join(f.home, 'skills', 'xlflow'))
  assert.equal(await readFile(join(result.directory, 'LICENSE'), 'utf8'), 'fixture')
  assert.deepEqual((await readdir(join(result.skillRoot, 'references'))).sort(), references.map(name => `${name}.md`).sort())
  assert.ok(f.calls.includes(`skill install --target ${join(f.home, 'skills')} --json`))
})

test('repeat setup validates but does not reinstall or overwrite user customizations', async t => {
  const f = await fixture(t)
  const first = await f.setup()
  await writeFile(join(first.skillRoot, 'references', 'forms.md'), 'user customization')
  f.calls.length = 0
  assert.deepEqual(await f.setup(), first)
  assert.deepEqual(f.calls, ['version --json', '--version-json'])
  assert.equal(await readFile(join(first.skillRoot, 'references', 'forms.md'), 'utf8'), 'user customization')
})

test('preexisting user skill collision is preserved and reported', async t => {
  const f = await fixture(t)
  const skill = join(f.home, 'skills', 'xlflow')
  await mkdir(skill, { recursive: true })
  await writeFile(join(skill, 'SKILL.md'), 'my skill')
  await assert.rejects(f.setup(), { code: 'SKILL_COLLISION' })
  assert.equal(await readFile(join(skill, 'SKILL.md'), 'utf8'), 'my skill')
  assert.ok(!f.calls.some(call => call.startsWith('skill install')))
})

test('failed installer, incomplete bundle and missing managed references never return success', async t => {
  for (const output of ['{"status":"error"}', '{"status":"ok"}', 'not JSON']) {
    const f = await fixture(t)
    const original = f.operations.run
    f.operations.run = async (command, args) => args[0] === 'skill' ? output : original(command, args)
    await assert.rejects(f.setup(), { code: output === 'not JSON' ? 'INVALID_RESPONSE'
      : output.includes('error') ? 'SKILL_INSTALL_FAILED' : 'SKILL_INCOMPLETE' })
  }
  const f = await fixture(t)
  const result = await f.setup()
  await rm(join(result.skillRoot, 'references', 'forms.md'))
  await assert.rejects(f.setup(), { code: 'SKILL_INCOMPLETE' })
})

test('interrupted or corrupt runtime is not replaced or accepted', async t => {
  const f = await fixture(t, { verifyFiles: undefined })
  delete f.operations.verifyFiles
  const directory = join(f.home, 'runtime', 'xlflow-0.35.0')
  await mkdir(directory, { recursive: true })
  await writeFile(join(directory, 'LICENSE'), 'leave this file')
  await assert.rejects(f.setup(), { code: 'RUNTIME_INVALID' })
  assert.deepEqual(f.calls, [])
  assert.equal(await readFile(join(directory, 'LICENSE'), 'utf8'), 'leave this file')
})

test('relative homes fail before IO', async () => {
  await assert.rejects(provisionXlflow('relative-home'), { code: 'HOME_REQUIRED' })
  await assert.rejects(provisionXlflow(dirname(fileURLToPath(import.meta.url))), { code: 'HOME_IN_CHECKOUT' })
})

test('rejected child operation cannot publish a runtime or leave staging residue', async t => {
  const f = await fixture(t, { async run() { throw new Error('child failed') } })
  await assert.rejects(f.setup(), /child failed/u)
  assert.deepEqual(await readdir(join(f.home, 'runtime')), [])
})
