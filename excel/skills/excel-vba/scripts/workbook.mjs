/** Guard xlflow's live-original workflow using pane metadata and read-only COM identity. */
import { execFile } from 'node:child_process'
import { createHash, randomUUID } from 'node:crypto'
import { lstat, mkdir, mkdtemp, readFile, readdir, rename, rm, writeFile } from 'node:fs/promises'
import { dirname, isAbsolute, join, resolve, win32 } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { promisify } from 'node:util'

const executeFile = promisify(execFile)
const scriptDirectory = dirname(fileURLToPath(import.meta.url))
const formats = new Map([['.xlsm', 52], ['.xlsb', 50], ['.xlam', 55], ['.xltm', 53]])
const sourceDirectories = ['modules', 'classes', 'forms', 'workbook']
const sourceCommands = new Set(['lint', 'analyze', 'architecture', 'impact', 'inspect-gui', 'fmt'])
const liveCommands = new Set(['push', 'test', 'run', 'macros', 'inspect', 'export-image', 'ui', 'form', 'list', 'save'])
const sourceInspections = new Set(['calls', 'symbols'])
const liveInspections = new Set(['workbook', 'sheets', 'range', 'cell', 'used-range', 'form'])

function sourceInspection(args) { return args[0] === 'inspect' && sourceInspections.has(args[1]) }
function liveCommand(args) {
  return args[0] === 'inspect' ? liveInspections.has(args[1])
    : liveCommands.has(args[0]) && !(args[0] === 'form' && args[1] === 'new')
}

function fail(code, message) { throw Object.assign(new Error(message), { code }) }
function normalize(path) { return win32.normalize(path).toLowerCase() }

async function optionalFile(path) {
  try { return await readFile(path, 'utf8') }
  catch (error) { if (error.code === 'ENOENT') return undefined; throw error }
}

async function run(command, args, cwd) {
  const env = Object.fromEntries(Object.entries(process.env).filter(([name]) => !/KEY|SECRET|TOKEN|PASSWORD/iu.test(name)))
  try {
    return { exitCode: 0, stdout: (await executeFile(command, args, {
      cwd, env, windowsHide: true, timeout: 120000, maxBuffer: 4 * 1024 * 1024,
    })).stdout }
  } catch (error) {
    if (typeof error.code === 'number' && error.stdout && !error.killed) {
      return { exitCode: error.code, stdout: error.stdout }
    }
    if (error.code === 'EPERM' || error.code === 'EACCES') {
      fail('PROCESS_ACCESS_DENIED', `Operating system denied child-process access (${error.code}). In a confined shell call, retry the exact helper command with per-call sandbox_permissions and a specific justification. This code does not identify the source of the restriction.`)
    }
    fail('PROCESS_UNAVAILABLE', 'Child process failed or timed out. Diagnose the failure before retrying; this result does not establish an access denial.')
  }
}

async function inventory() {
  if (process.platform !== 'win32') fail('WINDOWS_REQUIRED', 'Live workbook identity requires Windows Excel.')
  const powershell = join(process.env.SystemRoot ?? 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
  const code = await readFile(join(scriptDirectory, 'inventory.ps1'), 'utf8')
  const result = await run(powershell, ['-NoProfile', '-NonInteractive', '-EncodedCommand', Buffer.from(code, 'utf16le').toString('base64')])
  if (result.exitCode !== 0) fail('INVENTORY_UNAVAILABLE', 'Read-only Excel COM inventory failed. Dismiss blocking dialogs or retry the same helper call with the shell tool escalation after an actual sandbox denial.')
  try { return JSON.parse(result.stdout.replace(/^\uFEFF/u, '')) }
  catch (_error) { fail('INVENTORY_UNAVAILABLE', 'Excel inventory did not return JSON; no target was accepted.') }
}

async function verifyRuntime(home) {
  const directory = join(home, 'runtime', 'xlflow-0.35.0')
  for (const [name, hash] of [
    ['xlflow.exe', '98e0528188d37816b84dc448e4be2ef746a07d230f622094ddd616205a13a537'],
    ['xlflow-excel-bridge.exe', 'e202695121bf17091989ae827c83e1cd4ffd9e536ca78be5f81c741857e8dfc6'],
  ]) {
    let bytes
    try { bytes = await readFile(join(directory, name)) }
    catch (_error) { fail('RUNTIME_INVALID', 'The installed xlflow v0.35.0 runtime is unavailable. Repair the product installation; do not use PATH or install software from the agent.') }
    if (createHash('sha256').update(bytes).digest('hex') !== hash) {
      fail('RUNTIME_INVALID', 'The installed xlflow v0.35.0 runtime changed. Repair the product installation; do not use PATH or install software from the agent.')
    }
  }
  return join(directory, 'xlflow.exe')
}

/**
 * Validate a pane locator against a complete live inventory without writing any files.
 * @param {object} context Current request's pane metadata; its id is not a COM handle.
 * @param {object} operations Optional read-only inventory provider for focused tests.
 * @returns {Promise<object>} Unique live target, including unsaved state; rejects unsupported or uncertain identities.
 */
export async function validateBinding(context, operations = {}) {
  if (!context || typeof context.workbook?.id !== 'string' || !context.workbook.id
    || context.selection?.workbookId !== context.workbook.id
    || typeof context.worksheet?.id !== 'string' || typeof context.observedAt !== 'string'
    || !Number.isFinite(Date.parse(context.observedAt))) {
    fail('PANE_CONTEXT_REQUIRED', 'Supply the current request pane metadata. Request a new pane capture when location or current-pane identity is uncertain.')
  }
  const path = context.workbook.path
  if (typeof path !== 'string' || !path.trim()) {
    fail('LOCATOR_MISSING', 'The pane has no saved local workbook path. Save in an explicitly chosen macro-enabled format and submit fresh pane context; no conversion was performed.')
  }
  if (!/^[a-z]:[\\/]/iu.test(path) || /[\u0000-\u001f]/u.test(path)) {
    fail('LOCAL_PATH_REQUIRED', 'VBA binding requires an absolute local drive path. Cloud URLs, network paths and unsaved workbook names cannot be matched; obtain a local locator and fresh pane context.')
  }
  const absolute = win32.normalize(path)
  const format = formats.get(win32.extname(absolute).toLowerCase())
  if (!format) fail('MACRO_FORMAT_REQUIRED', 'Use an explicitly chosen .xlsm, .xlsb, .xlam or .xltm original. Ordinary cell work does not require VBA; no conversion or Save As was performed.')
  let file
  try { file = await lstat(absolute) }
  catch (error) {
    if (error.code !== 'ENOENT') throw error
    fail('TARGET_MISSING', 'The captured file no longer exists. Check Save As or relocation and submit fresh pane context.')
  }
  if (!file.isFile() || file.isSymbolicLink()) fail('TARGET_INVALID', 'The workbook locator must name a regular local file.')
  const state = await (operations.inventory ?? inventory)()
  if (!state || state.complete !== true || !Array.isArray(state.processes)
    || state.processes.some(row => !row || row.complete !== true || !Number.isInteger(row.pid) || row.pid <= 0
      || !Array.isArray(row.workbooks) || row.workbooks.some(book => !book || typeof book.path !== 'string'
        || !book.path || typeof book.directory !== 'string' || !Number.isInteger(book.format) || typeof book.saved !== 'boolean'))
    || new Set(state.processes.map(row => row.pid)).size !== state.processes.length) {
    fail('INVENTORY_INCOMPLETE', 'At least one Excel process or workbook could not be enumerated. Dismiss blocking dialogs, resolve access restrictions and retry; uniqueness is unknown.')
  }
  const matches = state.processes.flatMap(row => row.workbooks
    .filter(book => typeof book.path === 'string' && normalize(book.path) === normalize(absolute))
    .map(book => ({ ...book, pid: row.pid })))
  if (matches.length !== 1) {
    fail(matches.length ? 'TARGET_AMBIGUOUS' : 'TARGET_NOT_OPEN', matches.length
      ? 'The exact workbook is open more than once. Resolve the duplicate instances before retrying.'
      : 'No open Excel workbook matches the captured path. Check Save As or closure and submit fresh pane context; do not open a copied original.')
  }
  const match = matches[0]
  if (!match.directory || match.format !== format || typeof match.saved !== 'boolean') {
    fail('LIVE_FORMAT_MISMATCH', 'Live workbook location or format disagrees with the saved locator. Resolve Save As or format changes and submit fresh pane context.')
  }
  return { path: absolute, workbookId: context.workbook.id, pid: match.pid, saved: match.saved,
    observedAt: context.observedAt, paneIdentity: 'captured_locator_only' }
}

function configuration(target) {
  const quote = value => JSON.stringify(value)
  return `[excel]\npath = ${quote(normalize(target.path))}\nbridge = "dotnet"\n\n[src]\n${sourceDirectories.map(name => `${name} = "src/${name}"`).join('\n')}\n\n[userform]\ncode_source = "sidecar"\n`
}

async function checkProject(project, target) {
  if (await optionalFile(join(project, 'xlflow.toml')) !== configuration(target)) {
    fail('PROJECT_CONFLICT', 'Existing project configuration differs from the guarded original or source directories. Preserve it and resolve source authority explicitly.')
  }
  for (const path of [project, join(project, 'src'), ...sourceDirectories.map(name => join(project, 'src', name))]) {
    const info = await lstat(path)
    if (!info.isDirectory() || info.isSymbolicLink()) fail('PROJECT_CONFLICT', 'Managed source directories must remain real directories; no source was replaced.')
  }
  const binding = JSON.parse(await readFile(join(project, 'binding.json'), 'utf8'))
  if (normalize(binding.path) !== normalize(target.path) || binding.pid !== target.pid) {
    fail('BINDING_CHANGED', 'Workbook path or Excel process changed. Revalidate and explicitly rebind the project before resuming.')
  }
  return binding
}

async function sourcesEmpty(project) {
  for (const name of sourceDirectories) {
    if ((await readdir(join(project, 'src', name))).length) return false
  }
  return true
}

async function checkSession(project, target, required) {
  const text = await optionalFile(join(project, '.xlflow', 'session.json'))
  if (!text) {
    if (required) fail('SESSION_REQUIRED', 'Attach the validated original before running a live command.')
    return
  }
  const session = JSON.parse(text.replace(/^\uFEFF/u, ''))
  if (session.poisoned) fail('RECOVERY_REQUIRED', 'xlflow session is poisoned. Load the upstream recovery reference; do not save or blindly retry.')
  if (session.owner !== 'external' || typeof session.workbook_path !== 'string'
    || normalize(session.workbook_path) !== normalize(target.path) || session.pid !== target.pid) {
    fail('SESSION_CONFLICT', 'Existing xlflow session does not identify this external original and Excel process. Resolve it without replacing or stopping the user workbook.')
  }
}

function rejectWorkbookPosition(args) {
  // Among supported commands, only export-image accepts a workbook positional argument.
  if (args[0] !== 'export-image') return
  const values = new Set(['--sheet', '--range', '--out', '--output-dir', '--name', '--format', '--wait-timeout'])
  const switches = new Set(['--overwrite', '--session', '--json', '--wait', '--help', '-h'])
  for (let index = 1; index < args.length; index++) {
    const value = args[index]
    if (!value.startsWith('-') || value === '--') {
      fail('TARGET_OVERRIDE_REFUSED', 'Omit the export-image workbook positional argument; the configured original is already guarded.')
    }
    const flag = value.split('=', 1)[0]
    if (values.has(flag)) {
      if (!value.includes('=') && ++index >= args.length) fail('COMMAND_INVALID', `${flag} requires a value.`)
    } else if (!switches.has(flag)) {
      fail('COMMAND_INVALID', 'Use the pinned export-image help flags; unknown flag arity cannot establish an omitted workbook argument.')
    }
  }
}

function commandArguments(args) {
  if (!Array.isArray(args) || !args.length || args.some(value => typeof value !== 'string' || !value || value.includes('\0'))) {
    fail('COMMAND_INVALID', 'Pass xlflow arguments as a JSON array of strings.')
  }
  const session = args[0] === 'session' && ['attach', 'status', 'stop'].includes(args[1])
  if (!session && !sourceCommands.has(args[0]) && !liveCommands.has(args[0]) && !['status', 'pull'].includes(args[0])) {
    fail('COMMAND_UNSUPPORTED', 'This helper supports source checks and live-original commands. init, new, process management, session start, pack and file-backed operations are separate workflows.')
  }
  if (args.some(value => /^(?:--backend|--bridge|--workbook|--config|--project|--cwd|--active|--input|--save-as)(?:=|$)/u.test(value))) {
    fail('TARGET_OVERRIDE_REFUSED', 'Do not override the guarded project target, bridge or backend. Live pull uses the Excel backend automatically.')
  }
  if (args.some(value => /^(?:--session|--no-save)=/u.test(value))) {
    fail('TARGET_OVERRIDE_REFUSED', 'Do not negate session/no-save flags in the guarded workflow.')
  }
  rejectWorkbookPosition(args)
  if (args[0] === 'inspect' && !sourceInspections.has(args[1]) && !liveInspections.has(args[1])) {
    fail('COMMAND_UNSUPPORTED', 'Choose a supported source or workbook inspect subcommand.')
  }
  if (sourceInspection(args) && args.includes('--session')) {
    fail('COMMAND_INVALID', 'inspect calls/symbols read source and do not accept --session.')
  }
  if (args[0] === 'form' && !['new', 'build', 'snapshot', 'export-image'].includes(args[1])) {
    fail('COMMAND_UNSUPPORTED', 'Use form new, build, snapshot or export-image. Source layout migration requires a separate authority decision.')
  }
  if (['push', 'test'].includes(args[0]) && !args.includes('--no-save')) {
    fail('NO_SAVE_REQUIRED', 'Pass --no-save for push and test on the live original. Persist deliberately through save after verification.')
  }
  if (session && (args.length > 3 || args.slice(2).some(value => value !== '--json'))) {
    fail('COMMAND_INVALID', 'Use session attach/status/stop without lifecycle overrides; stop only detaches a validated external session.')
  }
  const result = args.filter(value => value !== '--json')
  if (args[0] === 'pull') result.push('--backend', 'excel')
  if ((liveCommand(args) || args[0] === 'pull') && !result.includes('--session')) result.push('--session')
  result.push('--json')
  return result
}

/**
 * Validate, prepare, rebind, bootstrap or execute a guarded xlflow command.
 * @param {object} request Action, current pane context, and command args when applicable.
 * @param {object} operations Home and optional inventory/process replacements for isolated tests.
 * @returns {Promise<object>} JSON-ready binding/project or authoritative xlflow result; rejects before a refused operation.
 */
export async function workbookRequest(request, operations = {}) {
  if (!['validate', 'prepare', 'rebind', 'bootstrap', 'command'].includes(request?.action)) fail('ACTION_INVALID', 'Choose validate, prepare, rebind, bootstrap or command.')
  const args = request.action === 'command' ? commandArguments(request.args) : undefined
  const target = await validateBinding(request.context, operations)
  if (request.action === 'validate') return { ok: true, target }
  const home = operations.home ?? resolve(scriptDirectory, '../../..')
  if (!isAbsolute(home)) fail('HOME_INVALID', 'The product must supply an absolute Excel Harness home.')
  const key = createHash('sha256').update(normalize(target.path)).digest('hex').slice(0, 24)
  const parent = join(home, 'vba-projects')
  const project = join(parent, key)
  const config = await optionalFile(join(project, 'xlflow.toml'))
  let created = false
  if (config === undefined && ['prepare', 'bootstrap'].includes(request.action)) {
    try {
      await lstat(project)
      fail('PROJECT_CONFLICT', 'Existing private project has no managed configuration. Preserve it and resolve the collision.')
    } catch (error) { if (error.code !== 'ENOENT') throw error }
    await mkdir(parent, { recursive: true })
    const stage = await mkdtemp(join(parent, '.prepare-'))
    try {
      for (const name of sourceDirectories) await mkdir(join(stage, 'src', name), { recursive: true })
      await writeFile(join(stage, 'xlflow.toml'), configuration(target), { flag: 'wx', mode: 0o600 })
      await writeFile(join(stage, 'binding.json'), `${JSON.stringify(target)}\n`, { flag: 'wx', mode: 0o600 })
      await rename(stage, project)
      created = true
    } finally {
      await rm(stage, { recursive: true, force: true })
    }
  }
  const binding = await checkProject(project, target).catch(error => {
    if (request.action !== 'rebind' || error.code !== 'BINDING_CHANGED') throw error
    return { workbookId: '' }
  })
  if (request.action === 'rebind') {
    await checkSession(project, target, false)
    const temporary = join(project, `.binding-${randomUUID()}.json`)
    try {
      await writeFile(temporary, `${JSON.stringify(target)}\n`, { flag: 'wx', mode: 0o600 })
      await rename(temporary, join(project, 'binding.json'))
    } finally { await rm(temporary, { force: true }) }
    return { ok: true, target, project, rebound: true }
  }
  if (binding.workbookId !== target.workbookId) fail('PANE_BINDING_CHANGED', 'The pane binding changed. Use rebind with the current request context to preserve and resume this project.')
  if (request.action === 'prepare') return { ok: true, target, project, created }
  if (request.action === 'bootstrap' && !created) {
    fail('PROJECT_EXISTS', 'Project already exists. Inspect source authority and session state, then use guarded commands; bootstrap never replaces source or a session.')
  }
  const cli = await (operations.verifyRuntime ?? verifyRuntime)(home)
  const invoke = async command => {
    const current = await validateBinding(request.context, operations)
    await checkProject(project, current)
    if (current.pid !== target.pid) fail('BINDING_CHANGED', 'Excel process changed before command execution; revalidate and rebind.')
    const attach = command[0] === 'session' && command[1] === 'attach'
    if (!sourceInspection(command)) {
      await checkSession(project, current, !attach && (liveCommand(command) || command[0] === 'pull' || (command[0] === 'session' && command[1] === 'stop')))
    }
    if (attach && await optionalFile(join(project, '.xlflow', 'session.json'))) fail('SESSION_EXISTS', 'A session already exists. Resume it through guarded commands; attach will not replace it.')
    if (command[0] === 'pull' && !await sourcesEmpty(project)) fail('SOURCE_EXISTS', 'Pull would replace existing source. Resolve source authority explicitly; this helper only pulls into empty source directories.')
    const result = await (operations.run ?? run)(cli, command, project)
    let output
    try { output = JSON.parse(result.stdout.replace(/^\uFEFF/u, '')) }
    catch (_error) { fail('CLI_OUTPUT_INVALID', 'xlflow returned no structured JSON. Inspect the command failure before retrying.') }
    return { ok: result.exitCode === 0 && output.status === 'ok' && output.recovery?.required !== true, exitCode: result.exitCode, output }
  }
  if (request.action === 'command') return { ...await invoke(args), target, project }
  const attached = await invoke(['session', 'attach', '--json'])
  if (!attached.ok) return { ...attached, target, project, phase: 'attach' }
  return { ...await invoke(['pull', '--backend', 'excel', '--session', '--json']), target, project, phase: 'pull' }
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  try {
    if (process.argv.length !== 3) fail('REQUEST_REQUIRED', 'Usage: node workbook.mjs <request.json>. Supply current pane context and structured command arguments.')
    const result = await workbookRequest(JSON.parse(await readFile(process.argv[2], 'utf8')))
    console.log(JSON.stringify(result))
    if (!result.ok) process.exitCode = 1
  } catch (error) {
    console.log(JSON.stringify({ ok: false, error: { code: error.code ?? 'HELPER_FAILED', message: error.message } }))
    process.exitCode = 1
  }
}
