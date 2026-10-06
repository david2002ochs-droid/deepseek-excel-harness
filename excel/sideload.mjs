/** Isolated development helper; each launch retains a uniquely named workbook for Excel. */
import { randomUUID } from 'node:crypto'
import { mkdtemp, rename } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import settings from 'office-addin-dev-settings'
import manifests from 'office-addin-manifest'
import open from 'open'

/** Register the current manifest and hand a fresh workbook to the desktop file association.
 * Run in a dedicated process: generation temporarily changes its temporary-directory environment.
 * Files remain available after handoff, including when opening fails; no existing workbooks are removed.
 * @param {string} manifestPath Manifest to register and embed in the workbook.
 * @param {object} operations Office registration, generation, and opening operations.
 * @returns {Promise<string>} Retained workbook path after the opening request succeeds.
 */
export async function sideload(manifestPath, operations = {
  readManifest: manifests.OfficeAddinManifest.readManifestFile,
  register: settings.registerAddIn,
  disableDebugging: settings.disableDebugging,
  disableLiveReload: settings.disableLiveReload,
  generate: settings.generateSideloadFile,
  open,
}) {
  const manifest = await operations.readManifest(manifestPath)
  if (!manifest.id) throw new Error('The Excel manifest requires an add-in ID.')
  await operations.register(manifestPath)
  await operations.disableDebugging(manifest.id)
  await operations.disableLiveReload(manifest.id)
  const directory = await mkdtemp(join(tmpdir(), 'harness-excel-'))
  const previous = { TEMP: process.env.TEMP, TMP: process.env.TMP, TMPDIR: process.env.TMPDIR }
  let generated
  try {
    for (const name of Object.keys(previous)) process.env[name] = directory
    generated = await operations.generate(manifests.OfficeApp.Excel, manifest)
  } finally {
    for (const [name, value] of Object.entries(previous)) {
      if (value === undefined) delete process.env[name]
      else process.env[name] = value
    }
  }
  if (dirname(resolve(generated)) !== directory) {
    throw new Error('Excel workbook generation returned a file outside its launch directory.')
  }
  // Excel distinguishes open workbooks by basename even when their directories differ.
  const workbook = join(directory, `Harness-${randomUUID()}.xlsx`)
  await rename(generated, workbook)
  await operations.open(workbook, { wait: false })
  return workbook
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    if (process.platform !== 'win32') throw new Error('Windows Excel Desktop is required.')
    await sideload(fileURLToPath(new URL('./manifest.xml', import.meta.url)))
  } catch (_error) {
    console.error('Excel sideload failed. Check Excel installation, manifest, and developer registration.')
    process.exitCode = 1
  }
}
