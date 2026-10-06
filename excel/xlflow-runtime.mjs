/** Private Windows xlflow installation; never resolves an executable through PATH. */
import { execFile } from 'node:child_process'
import { createHash } from 'node:crypto'
import { lstat, mkdir, mkdtemp, readFile, rename, rm, writeFile } from 'node:fs/promises'
import { dirname, isAbsolute, join, relative, resolve, sep } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const VERSION = '0.35.0'
const URL = `https://github.com/harumiWeb/xlflow/releases/download/v${VERSION}/xlflow_windows_x86_64.zip`
const ARCHIVE_HASH = 'aebee089c65b92b1d0424d069cc0183bc7854e19297364408869d5098599f2f7'
const FILE_HASHES = {
  'xlflow-excel-bridge.exe': 'e202695121bf17091989ae827c83e1cd4ffd9e536ca78be5f81c741857e8dfc6',
  'xlflow.exe': '98e0528188d37816b84dc448e4be2ef746a07d230f622094ddd616205a13a537',
  LICENSE: 'f58a3dc0eafdfbb9615856ba74f340c96db45d0f7423cbb28bbdc12bef5c08a1',
}
const SKILL_FILES = ['SKILL.md', ...[
  'code-analysis', 'debugging', 'forms', 'formulas', 'object-model-traps', 'pack',
  'recovery', 'testing', 'xlflow-ui',
].map(name => `references/${name}.md`)]

function fail(code, message) {
  return Object.assign(new Error(message), { code })
}

async function exists(path) {
  try { await lstat(path); return true }
  catch (error) { if (error.code === 'ENOENT') return false; throw error }
}

async function verifyArchive(path) {
  if (createHash('sha256').update(await readFile(path)).digest('hex') !== ARCHIVE_HASH) {
    throw fail('CHECKSUM_MISMATCH', 'xlflow download failed SHA256 verification; no executable was run.')
  }
}

async function verifyFiles(directory) {
  for (const [name, hash] of Object.entries(FILE_HASHES)) {
    const path = join(directory, name)
    if (!(await exists(path)) || !(await lstat(path)).isFile()
      || createHash('sha256').update(await readFile(path)).digest('hex') !== hash) {
      throw fail('RUNTIME_INVALID', `Pinned xlflow file is missing or changed: ${path}. Use a fresh Excel Harness home to reinstall.`)
    }
  }
}

async function download(path, signal) {
  const response = await fetch(URL, { signal: AbortSignal.any([AbortSignal.timeout(120000), ...(signal ? [signal] : [])]) })
  if (!response.ok) throw fail('DOWNLOAD_FAILED', `xlflow release download failed (HTTP ${response.status}).`)
  const chunks = []
  let length = 0
  for await (const chunk of response.body) {
    length += chunk.length
    if (length > 128 * 1024 * 1024) throw fail('DOWNLOAD_FAILED', 'xlflow release exceeds the download size limit.')
    chunks.push(chunk)
  }
  await writeFile(path, Buffer.concat(chunks), { flag: 'wx', mode: 0o600 })
}

async function run(command, args, cwd, signal) {
  const env = Object.fromEntries(Object.entries(process.env)
    .filter(([name]) => !/KEY|SECRET|TOKEN|PASSWORD/iu.test(name)))
  try {
    let child
    const result = new Promise(resolve => {
      child = execFile(command, args, {
        cwd, env, signal, windowsHide: true, timeout: 120000, maxBuffer: 1024 * 1024,
      }, (error, stdout) => resolve({ error, stdout }))
    })
    // Abort's callback can precede close. Keep ownership until the child and its pipes stop.
    const closed = new Promise(resolve => child.once('close', resolve))
    const outcome = await result
    await closed
    if (outcome.error) throw outcome.error
    return outcome.stdout
  } catch (error) {
    // Child diagnostics may contain private paths or inherited configuration; do not forward them.
    throw fail('PROCESS_FAILED', `xlflow setup command failed${error.killed ? ' or timed out' : ''}.`)
  }
}

async function extract(archive, directory, signal) {
  if (!process.env.SystemRoot) throw fail('WINDOWS_REQUIRED', 'Windows system directory is unavailable.')
  const quote = value => `'${value.replaceAll("'", "''")}'`
  await run(join(process.env.SystemRoot, 'System32/WindowsPowerShell/v1.0/powershell.exe'), [
    '-NoProfile', '-NonInteractive', '-Command',
    `$ErrorActionPreference = 'Stop'; Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::ExtractToDirectory(${quote(archive)}, ${quote(directory)})`,
  ], directory, signal)
}

function jsonResult(output, operation) {
  try { return JSON.parse(output) }
  catch (error) { throw fail('INVALID_RESPONSE', `xlflow ${operation} returned invalid JSON.`) }
}

async function validateVersions(directory, runner) {
  const cli = jsonResult(await runner(join(directory, 'xlflow.exe'), ['version', '--json'], directory), 'version')
  const bridge = jsonResult(await runner(join(directory, 'xlflow-excel-bridge.exe'), ['--version-json'], directory), 'bridge version')
  if (cli.status !== 'ok' || cli.version?.version !== VERSION
    || bridge.name !== 'xlflow-excel-bridge' || bridge.version !== '1.0.0' || bridge.protocol_version !== 1) {
    throw fail('VERSION_MISMATCH', 'xlflow CLI or bundled bridge version does not match the pinned release.')
  }
}

async function validateSkill(root) {
  for (const name of SKILL_FILES) {
    const path = join(root, name)
    if (!(await exists(path)) || !(await lstat(path)).isFile() || (await lstat(path)).size === 0) {
      throw fail('SKILL_INCOMPLETE', `xlflow skill file is missing or empty: ${path}. Restore this file without replacing your customizations.`)
    }
  }
}

/** Provision the pinned CLI, bridge, license and full upstream skill in a supplied Harness home.
 * Existing managed skill content is preserved; unowned collisions and corrupt installs fail without replacement.
 * Setup is Windows-only and each download/process has a two-minute deadline. Reuse rechecks executable hashes,
 * versions and skill completeness. Callers must serialize setup for the same home.
 * @param {string} home Absolute Excel Harness DSH_HOME outside the checkout.
 * @param {object} operations Optional AbortSignal plus test-only download, extraction, process and verification replacements.
 * @returns {Promise<{version:string, directory:string, cliPath:string, bridgePath:string, skillRoot:string}>} Verified private paths.
 */
export async function provisionXlflow(home, operations = {}) {
  const signal = operations.signal
  signal?.throwIfAborted()
  if (!isAbsolute(home)) throw fail('HOME_REQUIRED', 'Supply an absolute Excel Harness home directory.')
  home = resolve(home)
  const checkout = dirname(dirname(fileURLToPath(import.meta.url)))
  const fromCheckout = relative(checkout, home)
  if (fromCheckout === '' || (fromCheckout !== '..' && !fromCheckout.startsWith(`..${sep}`) && !isAbsolute(fromCheckout))) {
    throw fail('HOME_IN_CHECKOUT', 'The Excel Harness home must be outside the source checkout.')
  }
  if (process.platform !== 'win32' && !Object.entries(operations).some(([name, value]) => name !== 'signal' && typeof value === 'function')) {
    throw fail('WINDOWS_REQUIRED', 'The pinned xlflow runtime requires Windows.')
  }
  const replacements = { download, extract, run, verifyArchive, verifyFiles, ...operations }
  const io = Object.fromEntries(Object.entries(replacements).filter(([, value]) => typeof value === 'function').map(([name, operation]) => [name, async (...args) => {
    signal?.throwIfAborted()
    const result = await operation(...args, signal)
    signal?.throwIfAborted()
    return result
  }]))
  const directory = join(resolve(home), 'runtime', `xlflow-${VERSION}`)
  const skillParent = join(resolve(home), 'skills')
  const skillRoot = join(skillParent, 'xlflow')
  const receipt = join(directory, 'skill-install.json')
  if (!(await exists(directory))) {
    await mkdir(join(home, 'runtime'), { recursive: true })
    const staging = await mkdtemp(join(home, 'runtime', '.xlflow-'))
    if (dirname(staging) !== join(home, 'runtime')) throw fail('STAGING_INVALID', 'xlflow staging directory escaped the Harness runtime directory.')
    try {
      const archive = join(staging, 'release.zip')
      const unpacked = join(staging, 'unpacked')
      await mkdir(unpacked)
      await io.download(archive)
      await io.verifyArchive(archive)
      await io.extract(archive, unpacked)
      await io.verifyFiles(unpacked)
      await validateVersions(unpacked, io.run)
      await rename(unpacked, directory)
    } finally {
      // This randomly created directory is owned by this invocation, never a user-supplied cleanup target.
      await rm(staging, { recursive: true, force: true })
    }
  } else {
    await io.verifyFiles(directory)
    await validateVersions(directory, io.run)
  }
  if (await exists(skillRoot)) {
    if (!(await exists(receipt)) || (await readFile(receipt, 'utf8')) !== `${skillRoot}\n`) {
      throw fail('SKILL_COLLISION', `An unmanaged xlflow skill already exists at ${skillRoot}. Move it yourself before setup; it was not overwritten.`)
    }
  } else {
    signal?.throwIfAborted()
    await mkdir(skillParent, { recursive: true })
    const result = jsonResult(await io.run(join(directory, 'xlflow.exe'), [
      'skill', 'install', '--target', skillParent, '--json',
    ], directory), 'skill install')
    if (result.status !== 'ok') throw fail('SKILL_INSTALL_FAILED', 'xlflow skill installation did not succeed.')
    await validateSkill(skillRoot)
    await writeFile(receipt, `${skillRoot}\n`, { mode: 0o600 })
  }
  await validateSkill(skillRoot)
  signal?.throwIfAborted()
  return { version: VERSION, directory, cliPath: join(directory, 'xlflow.exe'),
    bridgePath: join(directory, 'xlflow-excel-bridge.exe'), skillRoot }
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  try {
    if (process.argv.length !== 4 || process.argv[2] !== '--home') {
      throw fail('USAGE', 'Usage: node excel/xlflow-runtime.mjs --home <absolute Excel Harness home>')
    }
    console.log(JSON.stringify({ status: 'ok', ...await provisionXlflow(process.argv[3]) }))
  } catch (error) {
    console.error(JSON.stringify({ status: 'error', code: error.code ?? 'SETUP_FAILED', message: error.message }))
    process.exitCode = 1
  }
}
