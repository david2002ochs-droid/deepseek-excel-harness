/** Owned subprocess lifecycle for the local Excel development entry. */
import { spawn } from 'node:child_process'
import { createServer } from 'node:net'
import { join } from 'node:path'

/** Reject an occupied loopback port without connecting to or stopping its owner.
 * @param {number} port Port required by the development entry.
 * @returns {Promise<void>} Resolves after the probe listener has closed.
 */
export async function requireFreePort(port) {
  const server = createServer()
  await new Promise((resolve, reject) => {
    server.once('error', () => reject(new Error(`Port ${port} is unavailable; close its owner yourself before starting Excel Harness.`)))
    server.listen(port, '127.0.0.1', resolve)
  })
  await new Promise((resolve, reject) => server.close(error => error ? reject(error) : resolve()))
}

/** Extract only the original Web profile's complete, validated startup line.
 * @param {string} output Private child output, never forwarded to the terminal.
 * @returns {string|undefined} Credential URL, retained only in memory.
 */
export function startupUrl(output) {
  return output.match(/(?:^|\n)dsh web: (https:\/\/localhost:3443\/\?token=[A-Za-z0-9_-]+)\r?\n/u)?.[1]
}

async function stopWindowsTree(pid) {
  if (!process.env.SystemRoot) throw new Error('Windows system directory is unavailable.')
  const killer = spawn(join(process.env.SystemRoot, 'System32', 'taskkill.exe'), ['/PID', String(pid), '/T', '/F'], {
    windowsHide: true, stdio: 'ignore', timeout: 10000,
  })
  await new Promise((resolve, reject) => {
    let failed = false
    killer.once('error', () => { failed = true })
    killer.once('close', code => failed || code !== 0
      ? reject(new Error('Windows process-tree termination failed.')) : resolve())
  })
}

function childProcess(spec, env, stopTree) {
  const child = spawn(spec.command, spec.args, {
    cwd: spec.cwd, env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'],
  })
  let closed = false
  let exited = false
  let output = ''
  const listeners = new Set()
  const collect = chunk => {
    output = (output + chunk.toString()).slice(-65536)
    for (const listener of listeners) listener()
  }
  child.stdout.on('data', collect)
  child.stderr.on('data', collect)
  child.once('exit', () => {
    exited = true
    // Descendants may inherit pipes. Their handles must not keep this child open.
    child.stdout.destroy()
    child.stderr.destroy()
  })
  const done = new Promise(resolve => {
    child.once('error', () => { /* close follows spawn failure; never echo credential-bearing diagnostics. */ })
    child.once('close', (code, signal) => {
      closed = true
      resolve({ code, signal })
      for (const listener of listeners) listener()
    })
  })
  return {
    child, done,
    get closed() { return closed },
    waitFor(match, signal, timeoutMs, label) {
      return new Promise((resolve, reject) => {
        const finish = (error, value) => {
          clearTimeout(timer)
          listeners.delete(check)
          signal.removeEventListener('abort', abort)
          if (error) reject(error)
          else resolve(value)
        }
        const abort = () => finish(new Error('Launch stopped.'))
        const check = () => {
          if (closed) return finish(new Error(`${label} exited before readiness.`))
          const value = match(output)
          if (value) finish(undefined, value)
        }
        const timer = setTimeout(() => finish(new Error(`${label} startup timed out.`)), timeoutMs)
        listeners.add(check)
        signal.addEventListener('abort', abort, { once: true })
        if (signal.aborted) abort()
        else check()
      })
    },
    async stop() {
      listeners.clear()
      output = ''
      if (closed) return
      if (exited || child.pid === undefined) { await done; return }
      let treeFailed = false
      try {
        // The PID comes only from this launch. Await the utility's close even on error.
        if (stopTree) await stopTree(child.pid)
      } catch (_error) {
        treeFailed = true
      } finally {
        // Helper failure must never skip direct-child termination and its awaited exit.
        if (!exited) child.kill(process.platform === 'win32' ? 'SIGKILL' : 'SIGTERM')
        const timer = setTimeout(() => { if (!exited) child.kill('SIGKILL') }, 10000)
        try { await done } finally { clearTimeout(timer) }
      }
      if (treeFailed) throw new Error('Process-tree cleanup failed; owned direct child stopped. Descendants may require attention.')
    },
  }
}

/** Start the profile, wrapper, and disposable-workbook sideload in order.
 * Child output stays private; only fixed stage messages reach report. Each stage
 * has a deadline. Failure or cancellation awaits termination of owned children.
 * @param {object} options Explicit commands, environment, cancellation and deadlines;
 * optional stopTree replaces the Windows cleanup helper for isolated failure tests.
 * @returns {Promise<void>} Resolves on cancellation; rejects on startup/runtime failure.
 */
export async function launch({ backend, wrapper, sideload, env, signal, timeoutMs, report, stopTree = process.platform === 'win32' ? stopWindowsTree : undefined }) {
  const children = []
  const controller = new AbortController()
  const cancel = () => controller.abort()
  signal.addEventListener('abort', cancel, { once: true })
  if (signal.aborted) cancel()
  const start = (spec, childEnv) => {
    if (controller.signal.aborted) throw new Error('Launch stopped.')
    const child = childProcess(spec, childEnv, stopTree)
    children.push(child)
    return child
  }
  let failure
  try {
    const host = start(backend, env)
    const unexpected = host.done.then(() => { throw new Error('Harness exited; Excel entry stopped.') })
    // Attach immediately so an exit during another stage cannot become unhandled.
    unexpected.catch(() => {})
    report('Starting Harness web profile…')
    const url = await host.waitFor(startupUrl, controller.signal, timeoutMs, 'Harness')
    const proxy = start(wrapper, { ...env, EXCEL_HARNESS_URL: url })
    const proxyExit = proxy.done.then(() => { throw new Error('HTTPS wrapper exited; Excel entry stopped.') })
    proxyExit.catch(() => {})
    report('Starting Excel HTTPS entry…')
    await Promise.race([
      proxy.waitFor(text => text.includes('Excel taskpane: https://localhost:3443/excel/taskpane.html\n'), controller.signal, timeoutMs, 'HTTPS wrapper'),
      unexpected,
    ])
    report('Opening a disposable Excel test workbook…')
    const office = start(sideload, env)
    let timer
    const deadline = new Promise((_, reject) => {
      timer = setTimeout(() => reject(new Error('Excel sideload timed out.')), timeoutMs)
    })
    let removeAbort
    const aborted = new Promise(resolve => {
      const handler = () => resolve('stopped')
      removeAbort = () => controller.signal.removeEventListener('abort', handler)
      controller.signal.addEventListener('abort', handler, { once: true })
      if (controller.signal.aborted) handler()
    })
    try {
      const result = await Promise.race([office.done, unexpected, proxyExit, deadline, aborted])
      if (result !== 'stopped') {
        if (result.code !== 0 || result.signal) throw new Error('Excel sideload failed. Check Excel installation and developer registration.')
        clearTimeout(timer)
        report('Sideload completed; local servers running. Keep this terminal open; press Ctrl+C to stop its servers.')
        await Promise.race([unexpected, proxyExit, aborted])
      }
    } finally {
      clearTimeout(timer)
      removeAbort()
    }
  } catch (error) {
    if (!signal.aborted) failure = error
  } finally {
    signal.removeEventListener('abort', cancel)
    controller.abort()
    const results = await Promise.allSettled(children.reverse().map(child => child.stop()))
    const cleanup = results.find(result => result.status === 'rejected')
    if (cleanup) failure = cleanup.reason
  }
  if (failure) throw failure
}
