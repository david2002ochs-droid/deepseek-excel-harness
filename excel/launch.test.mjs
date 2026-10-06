/** Fake children only: never import start.mjs or invoke the Harness, wrapper, or Excel. */
import assert from 'node:assert/strict'
import { spawn } from 'node:child_process'
import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { createServer } from 'node:net'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import test from 'node:test'
import { launch, requireFreePort, startupUrl } from './launch.mjs'

const fakeUrl = 'https://localhost:3443/?token=fake-test-only'
const source = `
const fs = require('node:fs');
const net = require('node:net');
const role = process.env.FAKE_ROLE;
if (role === 'wrapper' && process.env.EXCEL_HARNESS_URL !== '${fakeUrl}') process.exit(19);
if (role !== 'wrapper' && process.env.EXCEL_HARNESS_URL) process.exit(20);
fs.writeFileSync(role + '.pid', String(process.pid));
if (role === 'sideload') process.exit(Number(process.env.FAKE_FAIL || 0));
const server = net.createServer();
server.listen(0, '127.0.0.1', () => {
  fs.writeFileSync(role + '.port', String(server.address().port));
  if (process.env.FAKE_HANG === role) return;
  const text = role === 'backend' ? 'dsh web: ${fakeUrl}\\n' : 'Excel taskpane: https://localhost:3443/excel/taskpane.html\\n';
  process.stdout.write(text.slice(0, 12));
  setImmediate(() => process.stdout.write(text.slice(12)));
});
`

function fakeSpec(directory, role, extra = '') {
  return { command: process.execPath, args: ['-e', `process.env.FAKE_ROLE=${JSON.stringify(role)};${extra}${source}`], cwd: directory }
}

async function setup(t, options = {}) {
  const directory = await mkdtemp(join(tmpdir(), 'excel-launch-test-'))
  const controller = new AbortController()
  t.after(async () => {
    controller.abort()
    await rm(directory, { recursive: true, force: true })
  })
  const messages = []
  const configuration = {
    backend: fakeSpec(directory, 'backend'), wrapper: fakeSpec(directory, 'wrapper'),
    sideload: fakeSpec(directory, 'sideload'),
    env: { SystemRoot: process.env.SystemRoot, PATH: process.env.PATH },
    signal: controller.signal, timeoutMs: 10000,
    report(message) {
      messages.push(message)
      if (message.startsWith('Sideload completed; local servers running.')) controller.abort()
    },
    ...options,
  }
  return { directory, controller, messages, configuration }
}

async function assertStopped(directory, roles) {
  for (const role of roles) {
    const pid = Number(await readFile(join(directory, `${role}.pid`), 'utf8'))
    assert.throws(() => process.kill(pid, 0), { code: 'ESRCH' })
  }
}

test('accepts only the complete original startup line, including split output and CRLF', () => {
  assert.equal(startupUrl(`dsh web: ${fakeUrl}`), undefined)
  assert.equal(startupUrl(`noise\ndsh web: ${fakeUrl}\r\n`), fakeUrl)
  for (const url of [fakeUrl + '&extra=1', fakeUrl.replace('localhost', 'example.com'), fakeUrl.replace('https:', 'http:'), fakeUrl + '#fragment']) {
    assert.equal(startupUrl(`dsh web: ${url}\n`), undefined)
  }
})

test('orders fake stages, passes credential only to wrapper and awaits owned shutdown', async t => {
  const fixture = await setup(t)
  await launch(fixture.configuration)
  assert.equal(fixture.messages.length, 4)
  assert.ok(fixture.messages.every(message => !message.includes('fake-test-only')))
  await assertStopped(fixture.directory, ['backend', 'wrapper', 'sideload'])
})

test('sideload failure terminates both fake servers before rejection', async t => {
  const fixture = await setup(t)
  fixture.configuration.sideload = fakeSpec(fixture.directory, 'sideload', 'process.env.FAKE_FAIL="3";')
  await assert.rejects(launch(fixture.configuration), /Excel sideload failed/u)
  await assertStopped(fixture.directory, ['backend', 'wrapper', 'sideload'])
})

test('wrapper deadline cleans up owned fake children', async t => {
  const fixture = await setup(t)
  fixture.configuration.wrapper = fakeSpec(fixture.directory, 'wrapper', 'process.env.FAKE_HANG="wrapper";')
  fixture.configuration.timeoutMs = 2000
  await assert.rejects(launch(fixture.configuration), /HTTPS wrapper startup timed out/u)
  await assertStopped(fixture.directory, ['backend', 'wrapper'])
})

test('spawn failure is fixed diagnostic and stops the already-started fake backend', async t => {
  const fixture = await setup(t)
  fixture.configuration.wrapper = { command: join(fixture.directory, 'missing-executable'), args: [], cwd: fixture.directory }
  await assert.rejects(launch(fixture.configuration), /HTTPS wrapper exited before readiness|HTTPS wrapper exited; Excel entry stopped/u)
  await assertStopped(fixture.directory, ['backend'])
})

test('cancellation during wrapper startup prevents sideload and awaits both fake exits', async t => {
  const fixture = await setup(t)
  fixture.configuration.report = message => {
    if (message.startsWith('Starting Excel HTTPS entry')) fixture.controller.abort()
  }
  await launch(fixture.configuration)
  await assertStopped(fixture.directory, ['backend'])
  await assert.rejects(readFile(join(fixture.directory, 'sideload.pid')), { code: 'ENOENT' })
})

test('fake backend exit while the wrapper waits aborts startup and cleans up', async t => {
  const fixture = await setup(t)
  fixture.configuration.wrapper = fakeSpec(fixture.directory, 'wrapper', `setImmediate(() => process.kill(Number(require('node:fs').readFileSync('backend.pid'))));process.env.FAKE_HANG='wrapper';`)
  await assert.rejects(launch(fixture.configuration), /Harness exited/u)
  await assertStopped(fixture.directory, ['backend', 'wrapper'])
})

test('occupied ephemeral port is rejected without stopping its owner', async t => {
  const server = createServer()
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve))
  t.after(() => new Promise(resolve => server.close(resolve)))
  await assert.rejects(requireFreePort(server.address().port), /unavailable/u)
  assert.equal(server.listening, true)
})

test('cleanup-helper spawn failure still awaits actual owned child exits before rejecting', async t => {
  const fixture = await setup(t)
  let failures = 0
  fixture.configuration.stopTree = () => new Promise((resolve, reject) => {
    const helper = spawn(join(fixture.directory, 'missing-taskkill'), [], { stdio: 'ignore', windowsHide: true })
    let failed = false
    helper.once('error', () => { failed = true })
    helper.once('close', () => {
      if (failed) {
        failures++
        reject(new Error('private helper detail must not escape'))
      } else resolve()
    })
  })
  await assert.rejects(launch(fixture.configuration), {
    message: 'Process-tree cleanup failed; owned direct child stopped. Descendants may require attention.',
  })
  assert.equal(failures, 2)
  await assertStopped(fixture.directory, ['backend', 'wrapper', 'sideload'])
})
