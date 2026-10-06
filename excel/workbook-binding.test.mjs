/** Windows file-identity and sequencing checks; COM and xlflow operations are replaced. */
import assert from 'node:assert/strict'
import { mkdtemp, readFile, readdir, rm, writeFile, mkdir } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import test from 'node:test'
import { validateBinding, workbookRequest } from './skills/excel-vba/scripts/workbook.mjs'

async function fixture(t) {
  const root = await mkdtemp(join(tmpdir(), 'excel-vba-test-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const path = join(root, "Owner's draft.xlsm")
  await writeFile(path, 'original saved workbook bytes')
  const context = { workbook: { id: 'pane-1', path }, worksheet: { id: 'sheet-1', name: 'Sheet1' },
    selection: { workbookId: 'pane-1', address: 'Sheet1!A1' }, observedAt: new Date().toISOString() }
  const book = { path, directory: root, saved: false, format: 52 }
  let state = { complete: true, processes: [{ complete: true, pid: 123, workbooks: [book] }] }
  const calls = []
  let inventoryCount = 0
  const operations = { home: join(root, 'home'),
    inventory: async () => { inventoryCount++; return state },
    verifyRuntime: async home => join(home, 'runtime', 'xlflow-0.35.0', 'xlflow.exe'),
    run: async (cli, args, cwd) => {
      calls.push({ cli, args, cwd, inventoryCount })
      if (args[0] === 'session' && args[1] === 'attach') {
        await mkdir(join(cwd, '.xlflow'), { recursive: true })
        await writeFile(join(cwd, '.xlflow', 'session.json'), JSON.stringify({ pid: 123, owner: 'external', workbook_path: path }))
      }
      if (args[0] === 'pull') await writeFile(join(cwd, 'src', 'modules', 'Main.bas'), 'Option Explicit\n')
      return { exitCode: 0, stdout: JSON.stringify({ status: 'ok', command: args[0] }) }
    } }
  return { root, context, book, calls, operations, setState: value => { state = value }, request: action => workbookRequest({ action, context }, operations) }
}

const windows = { skip: process.platform !== 'win32' }

test('missing native inventory executable remains a generic failure without exposing child diagnostics', windows, async t => {
  const f = await fixture(t)
  const original = process.env.SystemRoot
  try {
    process.env.SystemRoot = f.root
    await assert.rejects(validateBinding(f.context), error => {
      assert.equal(error.code, 'PROCESS_UNAVAILABLE')
      assert.ok(!error.message.includes(f.root))
      assert.match(error.message, /does not establish an access denial/u)
      return true
    })
  } finally {
    if (original === undefined) delete process.env.SystemRoot
    else process.env.SystemRoot = original
  }
})

test('saved macro-enabled original with live unsaved edits validates without writing', windows, async t => {
  const f = await fixture(t)
  const result = await f.request('validate')
  assert.equal(result.target.saved, false)
  assert.equal(result.target.paneIdentity, 'captured_locator_only')
  assert.deepEqual(await readdir(f.root), ["Owner's draft.xlsm"])
  assert.equal(f.calls.length, 0)
})

test('duplicate exact workbook across instances refuses before private project creation', windows, async t => {
  const f = await fixture(t)
  f.setState({ complete: true, processes: [123, 456].map(pid => ({ complete: true, pid, workbooks: [f.book] })) })
  await assert.rejects(f.request('bootstrap'), { code: 'TARGET_AMBIGUOUS' })
  assert.deepEqual(await readdir(f.root), ["Owner's draft.xlsm"])
  assert.equal(f.calls.length, 0)
})

test('unreadable other Excel process cannot imply uniqueness', windows, async t => {
  const f = await fixture(t)
  f.setState({ complete: false, processes: [{ complete: true, pid: 123, workbooks: [f.book] }, { complete: false, pid: 456, workbooks: [] }] })
  await assert.rejects(f.request('prepare'), { code: 'INVENTORY_INCOMPLETE' })
  assert.equal(f.calls.length, 0)
})

test('malformed workbook inventory cannot be ignored when claiming uniqueness', windows, async t => {
  const f = await fixture(t)
  f.setState({ complete: true, processes: [{ complete: true, pid: 123, workbooks: [f.book, { path: '' }] }] })
  await assert.rejects(f.request('prepare'), { code: 'INVENTORY_INCOMPLETE' })
})

test('Save As live path change refuses the stale original even while old file exists', windows, async t => {
  const f = await fixture(t)
  f.setState({ complete: true, processes: [{ complete: true, pid: 123, workbooks: [{ ...f.book, path: join(f.root, 'Renamed.xlsm') }] }] })
  await assert.rejects(f.request('bootstrap'), { code: 'TARGET_NOT_OPEN' })
  assert.deepEqual(await readdir(f.root), ["Owner's draft.xlsm"])
})

for (const [label, path, code] of [
  ['missing', undefined, 'LOCATOR_MISSING'], ['cloud', 'https://example.test/Book.xlsm', 'LOCAL_PATH_REQUIRED'],
  ['unsaved name', 'Book1', 'LOCAL_PATH_REQUIRED'], ['network', '\\\\server\\Book.xlsm', 'LOCAL_PATH_REQUIRED'],
  ['xlsx', 'C:\\Book.xlsx', 'MACRO_FORMAT_REQUIRED'], ['vanished', 'C:\\missing-excel-vba-fixture.xlsm', 'TARGET_MISSING'],
]) {
  test(`${label} locator refuses before inventory or any write`, windows, async t => {
    const f = await fixture(t)
    f.context.workbook.path = path
    f.operations.inventory = async () => assert.fail('inventory must not run')
    await assert.rejects(f.request('bootstrap'), { code })
    assert.deepEqual(await readdir(f.root), ["Owner's draft.xlsm"])
  })
}

test('live format disagreement refuses without conversion', windows, async t => {
  const f = await fixture(t)
  f.book.format = 51
  await assert.rejects(f.request('prepare'), { code: 'LIVE_FORMAT_MISMATCH' })
})

test('prepare reuses a project without changing source, configuration or session', windows, async t => {
  const f = await fixture(t)
  const first = await f.request('prepare')
  const source = join(first.project, 'src', 'modules', 'Unsaved.bas')
  await writeFile(source, 'Private source changes\n')
  const config = await readFile(join(first.project, 'xlflow.toml'))
  const second = await f.request('prepare')
  assert.equal(second.created, false)
  assert.equal(second.project, first.project)
  assert.equal(await readFile(source, 'utf8'), 'Private source changes\n')
  assert.deepEqual(await readFile(join(first.project, 'xlflow.toml')), config)
  await assert.rejects(f.request('bootstrap'), { code: 'PROJECT_EXISTS' })
  assert.equal(f.calls.length, 0)
})

test('changed configuration fails without overwriting it', windows, async t => {
  const f = await fixture(t)
  const { project } = await f.request('prepare')
  await writeFile(join(project, 'xlflow.toml'), '[excel]\npath = "C:/Other.xlsm"\n')
  await assert.rejects(f.request('prepare'), { code: 'PROJECT_CONFLICT' })
  assert.equal(await readFile(join(project, 'xlflow.toml'), 'utf8'), '[excel]\npath = "C:/Other.xlsm"\n')
})

test('bootstrap revalidates independently before attach and forced live pull', windows, async t => {
  const f = await fixture(t)
  const result = await f.request('bootstrap')
  assert.equal(result.ok, true)
  assert.deepEqual(f.calls.map(call => call.args), [['session', 'attach', '--json'], ['pull', '--backend', 'excel', '--session', '--json']])
  assert.deepEqual(f.calls.map(call => call.inventoryCount), [2, 3])
  assert.equal(await readFile(f.context.workbook.path, 'utf8'), 'original saved workbook bytes')
  assert.equal(await readFile(join(result.project, 'src', 'modules', 'Main.bas'), 'utf8'), 'Option Explicit\n')
})

test('binding changed during bootstrap prevents attach', windows, async t => {
  const f = await fixture(t)
  let count = 0
  f.operations.inventory = async () => ({ complete: true, processes: [{ complete: true, pid: 123, workbooks: ++count === 1 ? [f.book] : [] }] })
  await assert.rejects(f.request('bootstrap'), { code: 'TARGET_NOT_OPEN' })
  assert.equal(f.calls.length, 0)
})

test('attach structured failure stops before pull', windows, async t => {
  const f = await fixture(t)
  f.operations.run = async (_cli, args) => { f.calls.push(args); return { exitCode: 1, stdout: '{"status":"failed","error":{"code":"vbide_access_denied"}}' } }
  const result = await f.request('bootstrap')
  assert.equal(result.ok, false)
  assert.equal(result.output.error.code, 'vbide_access_denied')
  assert.equal(result.phase, 'attach')
  assert.equal(f.calls.length, 1)
})

test('existing source and session are preserved on refused pull and attach', windows, async t => {
  const f = await fixture(t)
  const { project } = await f.request('bootstrap')
  const execute = args => workbookRequest({ action: 'command', context: f.context, args }, f.operations)
  await assert.rejects(execute(['pull']), { code: 'SOURCE_EXISTS' })
  await assert.rejects(execute(['session', 'attach']), { code: 'SESSION_EXISTS' })
  assert.equal(f.calls.length, 2)
  assert.equal(await readFile(join(project, 'src', 'modules', 'Main.bas'), 'utf8'), 'Option Explicit\n')
})

test('explicit same-workbook pane rebind resumes without replacing source', windows, async t => {
  const f = await fixture(t)
  const { project } = await f.request('bootstrap')
  f.context.workbook.id = 'new-pane'
  f.context.selection.workbookId = 'new-pane'
  await assert.rejects(f.request('prepare'), { code: 'PANE_BINDING_CHANGED' })
  const rebound = await f.request('rebind')
  assert.equal(rebound.rebound, true)
  assert.equal(rebound.project, project)
  const result = await workbookRequest({ action: 'command', context: f.context, args: ['push', '--fast', '--no-save'] }, f.operations)
  assert.equal(result.ok, true)
  assert.deepEqual(f.calls.at(-1).args, ['push', '--fast', '--no-save', '--session', '--json'])
  assert.equal(await readFile(join(project, 'src', 'modules', 'Main.bas'), 'utf8'), 'Option Explicit\n')
})

test('session contradiction and poison prevent command execution', windows, async t => {
  const f = await fixture(t)
  const { project } = await f.request('bootstrap')
  const session = join(project, '.xlflow', 'session.json')
  const request = () => workbookRequest({ action: 'command', context: f.context, args: ['save'] }, f.operations)
  await writeFile(session, JSON.stringify({ pid: 999, workbook_path: f.book.path, owner: 'managed' }))
  await assert.rejects(request(), { code: 'SESSION_CONFLICT' })
  await writeFile(session, JSON.stringify({ pid: 123, workbook_path: f.book.path, owner: 'external', poisoned: true }))
  await assert.rejects(request(), { code: 'RECOVERY_REQUIRED' })
  assert.equal(f.calls.length, 2)
})

test('guard refuses conversion, process lifecycle, target override and saving push before invocation', windows, async t => {
  const f = await fixture(t)
  for (const [args, code] of [[['init', f.book.path], 'COMMAND_UNSUPPORTED'], [['session', 'start'], 'COMMAND_UNSUPPORTED'],
    [['session', 'stop', '--discard'], 'COMMAND_INVALID'], [['push', '--backend=file'], 'TARGET_OVERRIDE_REFUSED'],
    [['push', '--no-save', '--no-save=false'], 'TARGET_OVERRIDE_REFUSED'], [['export-image', 'Other.xlsm'], 'TARGET_OVERRIDE_REFUSED'], [['push'], 'NO_SAVE_REQUIRED']]) {
    await assert.rejects(workbookRequest({ action: 'command', context: f.context, args }, f.operations), { code })
  }
  assert.equal(f.calls.length, 0)
})

test('run target override flags refuse split and equals xla/xlt values before inventory', windows, async t => {
  const f = await fixture(t)
  f.operations.inventory = async () => assert.fail('refused override must precede inventory')
  for (const flag of ['--input', '--save-as']) {
    for (const path of ['C:/Other.xla', 'C:/Other.xlt']) {
      for (const args of [['run', 'Main.Refresh', flag, path], ['run', 'Main.Refresh', `${flag}=${path}`]]) {
        await assert.rejects(workbookRequest({ action: 'command', context: f.context, args }, f.operations), { code: 'TARGET_OVERRIDE_REFUSED' })
      }
    }
  }
  assert.equal(f.calls.length, 0)
  assert.deepEqual(await readdir(f.root), ["Owner's draft.xlsm"])
})

test('export-image workbook position refuses any filename before inventory', windows, async t => {
  const f = await fixture(t)
  f.operations.inventory = async () => assert.fail('refused workbook position must precede inventory')
  for (const path of ['C:/Other.xla', 'C:/Other.xlt', 'extensionless-workbook']) {
    for (const args of [['export-image', path], ['export-image', '--sheet', 'Report', path], ['export-image', '--sheet=Report', '--', path]]) {
      await assert.rejects(workbookRequest({ action: 'command', context: f.context, args }, f.operations), { code: 'TARGET_OVERRIDE_REFUSED' })
    }
  }
  assert.equal(f.calls.length, 0)
})

test('run intentional save and workbook-looking macro argument stay on the original', windows, async t => {
  const f = await fixture(t)
  await f.request('bootstrap')
  await workbookRequest({ action: 'command', context: f.context, args: ['run', 'Main.Refresh', '--arg', 'string:C:/Other.xlsm', '--save'] }, f.operations)
  assert.deepEqual(f.calls.at(-1).args, ['run', 'Main.Refresh', '--arg', 'string:C:/Other.xlsm', '--save', '--session', '--json'])
  await workbookRequest({ action: 'command', context: f.context, args: ['export-image', '--sheet', 'Report', '--range=A1:D8', '--out', 'preview.png'] }, f.operations)
  assert.deepEqual(f.calls.at(-1).args, ['export-image', '--sheet', 'Report', '--range=A1:D8', '--out', 'preview.png', '--session', '--json'])
})

test('source inspection calls and symbols omit session checks and flags', windows, async t => {
  const f = await fixture(t)
  const { project } = await f.request('prepare')
  for (const args of [['inspect', 'calls', '--from', 'Main.Refresh'], ['inspect', 'symbols', '--path', 'src/modules/Main.bas']]) {
    const result = await workbookRequest({ action: 'command', context: f.context, args }, f.operations)
    assert.equal(result.ok, true)
    assert.deepEqual(f.calls.at(-1).args, [...args, '--json'])
  }
  await mkdir(join(project, '.xlflow'), { recursive: true })
  await writeFile(join(project, '.xlflow', 'session.json'), '{"poisoned":true}')
  const result = await workbookRequest({ action: 'command', context: f.context, args: ['inspect', 'calls'] }, f.operations)
  assert.equal(result.ok, true)
  assert.deepEqual(f.calls.at(-1).args, ['inspect', 'calls', '--json'])
})

test('live inspection workbook range and form still require and use external session', windows, async t => {
  const f = await fixture(t)
  await f.request('prepare')
  const cases = [['inspect', 'workbook'], ['inspect', 'range', 'Report!A1:D8'], ['inspect', 'form', 'Editor']]
  for (const args of cases) {
    await assert.rejects(workbookRequest({ action: 'command', context: f.context, args }, f.operations), { code: 'SESSION_REQUIRED' })
  }
  await workbookRequest({ action: 'command', context: f.context, args: ['session', 'attach'] }, f.operations)
  for (const args of cases) {
    await workbookRequest({ action: 'command', context: f.context, args }, f.operations)
    assert.deepEqual(f.calls.at(-1).args, [...args, '--session', '--json'])
  }
})

test('form new remains source-only while form build uses the guarded external session', windows, async t => {
  const f = await fixture(t)
  await f.request('prepare')
  await workbookRequest({ action: 'command', context: f.context, args: ['form', 'new', 'Editor'] }, f.operations)
  assert.deepEqual(f.calls.at(-1).args, ['form', 'new', 'Editor', '--json'])
  await assert.rejects(workbookRequest({ action: 'command', context: f.context, args: ['form', 'build', 'src/forms/Editor.yaml'] }, f.operations), { code: 'SESSION_REQUIRED' })
})

test('external stop detaches through pinned CLI without discard and recovery result stays authoritative', windows, async t => {
  const f = await fixture(t)
  await f.request('bootstrap')
  f.operations.run = async (_cli, args) => {
    assert.deepEqual(args, ['session', 'stop', '--json'])
    return { exitCode: 0, stdout: '{"status":"ok","recovery":{"required":true,"reason":"external_session_detached"}}' }
  }
  const result = await workbookRequest({ action: 'command', context: f.context, args: ['session', 'stop'] }, f.operations)
  assert.equal(result.ok, false)
  assert.equal(result.output.recovery.reason, 'external_session_detached')
})

test('selection workbook mismatch rejects invalid pane metadata', windows, async t => {
  const f = await fixture(t)
  f.context.selection.workbookId = 'other-pane'
  await assert.rejects(validateBinding(f.context, f.operations), { code: 'PANE_CONTEXT_REQUIRED' })
})
