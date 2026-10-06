/** Serialized private Excel-home setup; no server, workbook or settings changes. */
import { cp, lstat, mkdir, mkdtemp, readFile, realpath, rename, rm, rmdir, writeFile } from 'node:fs/promises'
import { delimiter, dirname, isAbsolute, join, relative, resolve, sep } from 'node:path'
import { setTimeout as delay } from 'node:timers/promises'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { provisionXlflow } from './xlflow-runtime.mjs'

const directory = dirname(fileURLToPath(import.meta.url))
const checkout = dirname(directory)
const instruction = '<!-- Excel VBA managed instruction -->\nFor VBA work requested from the Excel pane, load the excel-vba skill before xlflow and use its guarded original-workbook workflow. Ordinary cell work does not require VBA initialization.\n<!-- End Excel VBA managed instruction -->\n'
const required = ['SKILL.md', 'scripts/workbook.mjs', 'scripts/inventory.ps1', 'README.md', 'README.zh.md', 'README.i18n.yaml']

function fail(code, message) { return Object.assign(new Error(message), { code }) }

async function exists(path) {
  try { await lstat(path); return true }
  catch (error) { if (error.code === 'ENOENT') return false; throw error }
}

/** Reject unavailable drive roots and Excel homes inside this checkout, including linked ancestors.
 * @param {string} home Absolute home outside the source checkout.
 * @param {AbortSignal} [signal] Cancellation observed between filesystem lookups.
 * @returns {Promise<string>} Resolved safe home location.
 */
export async function validateExcelHome(home, signal) {
  signal?.throwIfAborted()
  if (!isAbsolute(home)) throw fail('HOME_REQUIRED', 'Supply an absolute Excel Harness home.')
  home = resolve(home)
  let ancestor = home
  while (true) {
    signal?.throwIfAborted()
    try { ancestor = await realpath(ancestor); break }
    catch (error) {
      signal?.throwIfAborted()
      if (error.code !== 'ENOENT') throw error
      const parent = dirname(ancestor)
      if (parent === ancestor) throw fail('HOME_UNAVAILABLE', 'The Excel Harness home drive is unavailable. Choose an available external directory.')
      ancestor = parent
    }
  }
  signal?.throwIfAborted()
  const root = await realpath(checkout)
  signal?.throwIfAborted()
  const inside = path => {
    const location = relative(root, path)
    return !isAbsolute(location) && location !== '..' && !location.startsWith(`..${sep}`)
  }
  if (inside(home) || inside(ancestor)) throw fail('HOME_IN_CHECKOUT', 'The Excel Harness home must be outside the source checkout.')
  return home
}

async function installSkill(home, signal) {
  const parent = join(home, 'skills')
  const target = join(parent, 'excel-vba')
  const receipt = join(home, 'runtime', 'excel-vba-install.txt')
  if (await exists(target)) {
    if (!(await exists(receipt)) || await readFile(receipt, 'utf8') !== `${target}\n`) {
      throw fail('SKILL_COLLISION', 'An unmanaged excel-vba skill exists; move it yourself before setup. It was not overwritten.')
    }
  } else {
    await mkdir(parent, { recursive: true })
    const staging = await mkdtemp(join(parent, '.excel-vba-'))
    try {
      await cp(join(directory, 'skills', 'excel-vba'), join(staging, 'skill'), { recursive: true, dereference: false })
      signal.throwIfAborted()
      await rename(join(staging, 'skill'), target)
      await writeFile(receipt, `${target}\n`, { flag: 'wx', mode: 0o600 })
    } finally {
      // The randomly created staging directory belongs only to this invocation.
      await rm(staging, { recursive: true, force: true })
    }
  }
  for (const file of required) {
    const path = join(target, file)
    if (!(await exists(path)) || !(await lstat(path)).isFile() || (await lstat(path)).size === 0) {
      throw fail('SKILL_INCOMPLETE', 'An excel-vba resource is missing or empty; restore it without replacing customizations.')
    }
  }
  return target
}

async function installInstruction(home, signal) {
  const path = join(home, 'AGENTS.md')
  let original = ''
  if (await exists(path)) original = await readFile(path, 'utf8')
  if (original.includes('<!-- Excel VBA managed instruction -->') || original.includes('<!-- End Excel VBA managed instruction -->')) {
    if (!original.includes(instruction)) throw fail('INSTRUCTION_CHANGED', 'The managed Excel VBA instruction was edited; reconcile it before setup. Existing instructions were preserved.')
    return
  }
  signal.throwIfAborted()
  // Append only: the existing user's bytes and instructions remain intact.
  await writeFile(path, `${original && !original.endsWith('\n') ? '\n' : ''}${original ? '\n' : ''}${instruction}`, { flag: 'a', mode: 0o600 })
}

/** Prepare a pinned runtime and both skills, preserving user skills and instructions.
 * One filesystem lock serializes same-home launchers; a two-minute deadline bounds waiting and setup.
 * An abandoned lock is refused at the deadline and requires manual repair after its owner has stopped.
 * Cancellation awaits native setup children and removes only this invocation's lock/staging directories.
 * @param {string} home Absolute Excel Harness home outside the checkout.
 * @param {object} options Optional signal; provision and timeoutMs replace operations only for tests.
 * @returns {Promise<object>} Verified runtime paths and installed skill directory.
 */
export async function setupExcelHome(home, options = {}) {
  const signal = AbortSignal.any([AbortSignal.timeout(options.timeoutMs ?? 120000), ...(options.signal ? [options.signal] : [])])
  signal.throwIfAborted()
  home = await validateExcelHome(home, signal)
  signal.throwIfAborted()
  await mkdir(home, { recursive: true, mode: 0o700 })
  const lock = join(home, '.excel-setup.lock')
  let owned = false
  try {
    while (!owned) {
      signal.throwIfAborted()
      try { await mkdir(lock); owned = true }
      catch (error) { if (error.code !== 'EEXIST') throw error; await delay(50, undefined, { signal }) }
    }
    signal.throwIfAborted()
    const runtime = await (options.provision ?? provisionXlflow)(home, { signal })
    signal.throwIfAborted()
    const skillRoot = await installSkill(home, signal)
    signal.throwIfAborted()
    await installInstruction(home, signal)
    signal.throwIfAborted()
    return { ...runtime, excelVbaSkillRoot: skillRoot }
  } finally {
    if (owned) await rmdir(lock)
  }
}

/** Add the verified runtime to a child environment without changing other values.
 * @param {object} env Environment already scrubbed by the launcher.
 * @param {object} runtime Successful setup result.
 * @returns {object} Copy with the pinned runtime first on the existing PATH.
 */
export function runtimeEnvironment(env, runtime) {
  const name = Object.keys(env).find(key => key.toUpperCase() === 'PATH') ?? 'PATH'
  return { ...env, [name]: `${runtime.directory}${env[name] ? `${delimiter}${env[name]}` : ''}` }
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const controller = new AbortController()
  const stop = () => controller.abort()
  process.on('SIGINT', stop)
  process.on('SIGTERM', stop)
  try {
    const home = process.argv.length === 4 && process.argv[2] === '--home' ? process.argv[3]
      : process.argv.length === 2 ? process.env.EXCEL_DSH_HOME ?? (process.env.LOCALAPPDATA && join(process.env.LOCALAPPDATA, 'DeepSeekHarnessExcel', 'Harness')) : undefined
    if (!home) throw fail('USAGE', 'Usage: node excel/setup.mjs [--home <absolute Excel Harness home>]')
    console.log(JSON.stringify({ status: 'ok', ...await setupExcelHome(home, { signal: controller.signal }) }))
  } catch (error) {
    console.error(JSON.stringify({ status: 'error', code: error.code ?? 'SETUP_FAILED', message: 'Excel home setup failed. Check the home, managed resources and setup lock.' }))
    process.exitCode = 1
  } finally {
    process.removeListener('SIGINT', stop)
    process.removeListener('SIGTERM', stop)
  }
}
