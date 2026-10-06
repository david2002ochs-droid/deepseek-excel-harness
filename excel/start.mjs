/** One-command development launch; no installer or login integration. */
import { access, mkdir, realpath } from 'node:fs/promises'
import { homedir } from 'node:os'
import { dirname, isAbsolute, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { launch, requireFreePort } from './launch.mjs'

const directory = dirname(fileURLToPath(import.meta.url))
const root = dirname(directory)
const controller = new AbortController()
const stop = () => controller.abort()
process.on('SIGINT', stop)
process.on('SIGTERM', stop)

try {
  if (process.platform !== 'win32') throw new Error('This development entry requires Windows Excel Desktop.')
  const npm = process.env.npm_execpath
  if (!npm) throw new Error('Start with npm --prefix excel start.')
  if (!process.env.EXCEL_DSH_HOME && !process.env.LOCALAPPDATA) {
    throw new Error('LOCALAPPDATA is unavailable. Set EXCEL_DSH_HOME to a directory outside the checkout.')
  }
  const home = resolve(process.env.EXCEL_DSH_HOME ?? join(process.env.LOCALAPPDATA, 'DeepSeekHarnessExcel', 'Harness'))
  const checkout = await realpath(root)
  let ancestor = home
  while (true) {
    try { ancestor = await realpath(ancestor); break } catch (error) {
      if (error.code !== 'ENOENT') throw error
      ancestor = dirname(ancestor)
    }
  }
  const isInside = path => {
    const location = relative(checkout, path)
    return !isAbsolute(location) && location !== '..' && !location.startsWith('..\\') && !location.startsWith('../')
  }
  if (isInside(home) || isInside(ancestor)) {
    throw new Error('EXCEL_DSH_HOME must be outside the checkout.')
  }
  await mkdir(home, { recursive: true, mode: 0o700 })
  const certificates = process.env.EXCEL_TLS_DIRECTORY ?? join(homedir(), '.office-addin-dev-certs')
  try {
    await access(join(certificates, 'localhost.key'))
    await access(join(certificates, 'localhost.crt'))
  } catch (_error) {
    throw new Error('Development certificates missing. Run npm --prefix excel run certs once.')
  }
  await requireFreePort(3080)
  await requireFreePort(3443)
  const env = Object.fromEntries(Object.entries(process.env).filter(([key]) =>
    !/KEY|SECRET|TOKEN|PASSWORD/iu.test(key) && !/^EXCEL_HARNESS_URL$/iu.test(key)))
  env.DSH_HOME = home
  await launch({
    // Same supported dsh CLI/profile entry used by the repository's dsh script.
    backend: { command: process.execPath, args: ['--import', 'tsx/esm', 'apps/cli/src/bin.ts', '--profile', 'web', '--no-open', '--host', '127.0.0.1', '--port', '3080', '--public-url', 'https://localhost:3443', '--trusted-host', 'localhost:3443'], cwd: root },
    wrapper: { command: process.execPath, args: [join(directory, 'serve.mjs')], cwd: directory },
    sideload: { command: process.execPath, args: [npm, 'exec', '--yes', '--package=office-addin-debugging@7.0.1', '--', 'office-addin-debugging', 'start', 'manifest.xml', 'desktop', '--app', 'excel', '--no-debug', '--no-live-reload', '--dev-server-port', '3443'], cwd: directory },
    env, signal: controller.signal, timeoutMs: 120000,
    report: message => console.log(message),
  })
} catch (error) {
  // Only our fixed messages are safe to show; filesystem errors can contain private paths.
  const message = error instanceof Error && !('code' in error) ? error.message : 'Excel entry setup failed. Check prerequisites and directory access.'
  console.error(message)
  process.exitCode = 1
} finally {
  process.removeListener('SIGINT', stop)
  process.removeListener('SIGTERM', stop)
}
