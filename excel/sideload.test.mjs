/** Real Office workbook generation with registration and desktop opening replaced; no Excel launches. */
import assert from 'node:assert/strict'
import { readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import test from 'node:test'
import settings from 'office-addin-dev-settings'
import manifests from 'office-addin-manifest'
import { sideload } from './sideload.mjs'

const manifestPath = fileURLToPath(new URL('./manifest.xml', import.meta.url))

function setup(t) {
  const root = resolve(tmpdir())
  const directories = new Set()
  const calls = []
  t.after(async () => {
    for (const directory of directories) {
      assert.equal(dirname(resolve(directory)), root)
      assert.match(basename(directory), /^harness-excel-/u)
      await rm(directory, { recursive: true, force: true })
    }
  })
  const operations = {
    readManifest: manifests.OfficeAddinManifest.readManifestFile,
    async register(path) { calls.push(['register', path]) },
    async disableDebugging(id) { calls.push(['debugging', id]) },
    async disableLiveReload(id) { calls.push(['reload', id]) },
    async generate(app, manifest) {
      directories.add(tmpdir())
      return settings.generateSideloadFile(app, manifest)
    },
    async open(path, options) { calls.push(['open', path, options]) },
  }
  return { operations, calls, directories }
}

test('repeated launches retain distinct workbook names and use the current manifest registration', async t => {
  const { operations, calls } = setup(t)
  const environment = { TEMP: process.env.TEMP, TMP: process.env.TMP, TMPDIR: process.env.TMPDIR }
  const first = await sideload(manifestPath, operations)
  const firstBytes = await readFile(first)
  const second = await sideload(manifestPath, operations)
  assert.notEqual(basename(first), basename(second))
  assert.notEqual(dirname(first), dirname(second))
  assert.deepEqual(await readFile(first), firstBytes)
  assert.equal((await readFile(second)).subarray(0, 2).toString(), 'PK')
  const manifest = await manifests.OfficeAddinManifest.readManifestFile(manifestPath)
  for (const [index, workbook] of [first, second].entries()) {
    assert.deepEqual(calls.slice(index * 4, index * 4 + 4), [
      ['register', manifestPath], ['debugging', manifest.id], ['reload', manifest.id],
      ['open', workbook, { wait: false }],
    ])
  }
  assert.deepEqual({ TEMP: process.env.TEMP, TMP: process.env.TMP, TMPDIR: process.env.TMPDIR }, environment)
})

test('registration failure stops before generation or opening', async t => {
  const { operations, calls, directories } = setup(t)
  operations.register = async () => { throw new Error('registration rejected') }
  await assert.rejects(sideload(manifestPath, operations), /registration rejected/u)
  assert.deepEqual(calls, [])
  assert.equal(directories.size, 0)
})

test('a manifest without an add-in ID is rejected before registration', async t => {
  const { operations, calls, directories } = setup(t)
  operations.readManifest = async () => ({})
  await assert.rejects(sideload(manifestPath, operations), /requires an add-in ID/u)
  assert.deepEqual(calls, [])
  assert.equal(directories.size, 0)
})

test('generation failure restores the environment and does not open Excel', async t => {
  const { operations, calls, directories } = setup(t)
  const environment = { TEMP: process.env.TEMP, TMP: process.env.TMP, TMPDIR: process.env.TMPDIR }
  operations.generate = async () => {
    directories.add(tmpdir())
    throw new Error('generation rejected')
  }
  await assert.rejects(sideload(manifestPath, operations), /generation rejected/u)
  assert.deepEqual({ TEMP: process.env.TEMP, TMP: process.env.TMP, TMPDIR: process.env.TMPDIR }, environment)
  assert.equal(calls.some(([operation]) => operation === 'open'), false)
})

test('a generated path outside the owned directory is never renamed or opened', async t => {
  const { operations, calls, directories } = setup(t)
  const original = await readFile(manifestPath)
  operations.generate = async () => {
    directories.add(tmpdir())
    return manifestPath
  }
  await assert.rejects(sideload(manifestPath, operations), /outside its launch directory/u)
  assert.deepEqual(await readFile(manifestPath), original)
  assert.equal(calls.some(([operation]) => operation === 'open'), false)
})

test('an opening failure propagates while retaining the workbook', async t => {
  const { operations } = setup(t)
  let workbook
  operations.open = async path => {
    workbook = path
    throw new Error('opening rejected')
  }
  await assert.rejects(sideload(manifestPath, operations), /opening rejected/u)
  assert.equal((await readFile(workbook)).subarray(0, 2).toString(), 'PK')
})
