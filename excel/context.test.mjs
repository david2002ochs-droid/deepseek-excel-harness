import assert from 'node:assert/strict'
import test from 'node:test'
import { captureExcelContext, cellCitationTarget, installExcelContextBridge, navigateCellCitation, workbookLocation } from './context.mjs'

function officeHost({ address = "'Modell Ü'!B5:D12", multi = true, supportsName = true,
  name = 'Test Book.xlsm', url = 'file:///C:/Models/Test%20Book.xlsm' } = {}) {
  const reads = []
  const worksheet = { id: 'sheet-1', name: 'Modell Ü', load: properties => reads.push(['worksheet', properties]) }
  const selection = { address, load: properties => reads.push(['selection', properties]) }
  for (const key of ['values', 'formulas', 'text']) Object.defineProperty(selection, key, {
    get() { throw new Error(`Unexpected cell-content read: ${key}`) },
  })
  return { reads,
    office: { context: { document: { url }, requirements: { isSetSupported: (_api, version) => version === '1.7' ? supportsName : multi } } },
    excel: { run: async action => action({ workbook: {
      name, load: properties => reads.push(['workbook', properties]),
      worksheets: { getActiveWorksheet: () => worksheet },
      getSelectedRanges: () => { reads.push(['multi']); return selection },
      getSelectedRange: () => { reads.push(['single']); return selection },
    }, sync: async () => { reads.push(['sync']) } }) },
  }
}

test('captures qualified selection and stable pane identity without reading cells or changing a workbook', async () => {
  const host = officeHost()
  const result = await captureExcelContext({ ...host, workbookId: 'pane-1', now: () => new Date('2026-10-06T10:00:00Z') })
  assert.deepEqual(result, { workbook: { id: 'pane-1', name: 'Test Book.xlsm', path: 'C:\\Models\\Test Book.xlsm' },
    worksheet: { id: 'sheet-1', name: 'Modell Ü' }, selection: { workbookId: 'pane-1', address: "'Modell Ü'!B5:D12" },
    observedAt: '2026-10-06T10:00:00.000Z' })
  assert.deepEqual(host.reads, [['workbook', 'name'], ['worksheet', 'id,name'], ['multi'], ['selection', 'address'], ['sync']])
})

test('preserves multiple selected areas and falls back to single-range capture only for older hosts', async () => {
  const areas = 'Sheet1!A1:B2,Sheet1!D5:E7'
  assert.equal((await captureExcelContext({ ...officeHost({ address: areas }), workbookId: 'pane-1' })).selection.address, areas)
  const host = officeHost({ multi: false, supportsName: false, url: '' })
  const result = await captureExcelContext({ ...host, workbookId: 'pane-1' })
  assert.deepEqual(result.workbook, { id: 'pane-1' })
  assert.ok(host.reads.some(read => read[0] === 'single'))
})

test('identifies unsaved and cloud workbooks by the host name rather than a URL basename', async () => {
  for (const url of ['', 'https://example.com/WopiFrame.aspx?token=secret']) {
    const result = await captureExcelContext({ ...officeHost({ name: 'Budget 2027.xlsx', url }), workbookId: 'pane-1' })
    assert.equal(result.workbook.name, 'Budget 2027.xlsx')
    assert.equal(result.workbook.path, url ? 'https://example.com/WopiFrame.aspx' : undefined)
  }
  const legacy = await captureExcelContext({ ...officeHost({ supportsName: false }), workbookId: 'pane-1' })
  assert.equal(legacy.workbook.name, 'Test Book.xlsm')
  await assert.rejects(captureExcelContext({ ...officeHost({ name: '' }), workbookId: 'pane-1' }), { message: 'context_unavailable' })
})

test('keeps native paths and strips web credentials, query strings and fragments from locators', () => {
  assert.equal(workbookLocation('C:\\Models\\Book.xlsm'), 'C:\\Models\\Book.xlsm')
  assert.equal(workbookLocation('file://server/share/Book.xlsm'), '\\\\server\\share\\Book.xlsm')
  assert.equal(workbookLocation('https://user:secret@example.com/Book.xlsm?token=secret#secret'), 'https://example.com/Book.xlsm')
  for (const value of [null, '', 'not a URL', 'javascript:alert(1)', 'file:///bad%ZZ']) assert.equal(workbookLocation(value), undefined)
})

test('rejects oversized selections rather than truncating a target or loading its contents', async () => {
  await assert.rejects(captureExcelContext({ ...officeHost({ address: 'Ü'.repeat(1200) }), workbookId: 'pane-1' }),
    { message: 'context_too_large' })
})

function bridgeHost(capture, navigate) {
  let listener
  let ids = 0
  const office = officeHost().office
  const replies = []
  const child = { postMessage: (data, origin) => replies.push({ data, origin }) }
  const window = { crypto: { randomUUID: () => `pane-${++ids}` },
    addEventListener: (_type, handler) => { listener = handler },
    removeEventListener: (_type, handler) => { assert.equal(handler, listener); listener = undefined },
  }
  const origin = 'https://localhost:3443'
  const dispose = installExcelContextBridge({ window, frame: { contentWindow: child }, origin, office, capture, navigate })
  return { replies, dispose, office, send: (data, overrides = {}) => listener?.({ data, origin, source: child, ...overrides }) }
}

const request = { type: 'dsh/excel-context/request', version: 1, requestId: 'request-1' }

test('ignores another origin, another frame and malformed requests before any Office read', async () => {
  let calls = 0
  const host = bridgeHost(async () => { calls++; return {} })
  await host.send(request, { origin: 'https://example.com' })
  await host.send(request, { source: {} })
  for (const data of [null, { ...request, version: 2 }, { ...request, requestId: 'x'.repeat(129) },
    { ...request, requestId: '<bad>' }, { ...request, type: 'other' }]) await host.send(data)
  assert.equal(calls, 0)
  assert.deepEqual(host.replies, [])
  host.dispose()
})

test('returns only fixed errors and correlates capture with its original request', async () => {
  const host = bridgeHost(async () => { throw new Error('private workbook path and secret') })
  await host.send(request)
  assert.deepEqual(host.replies, [{ origin: 'https://localhost:3443', data: {
    type: 'dsh/excel-context/response', version: 1, requestId: 'request-1', error: 'context_unavailable',
  } }])
  host.dispose()
})

test('does not queue a later selection behind a pending capture or publish after disposal', async () => {
  let finish
  const host = bridgeHost(() => new Promise(resolve => { finish = resolve }))
  const first = host.send(request)
  await host.send({ ...request, requestId: 'request-2' })
  assert.equal(host.replies[0].data.requestId, 'request-2')
  assert.equal(host.replies[0].data.error, 'context_unavailable')
  host.dispose()
  finish({ workbook: { id: 'pane-1' } })
  await first
  assert.equal(host.replies.length, 1)
})

test('accepts quoted localized cell citations and rejects external, unbounded and invalid targets', () => {
  assert.deepEqual(cellCitationTarget("'Café O''Brien'!$b$5:D12"), { sheet: "Café O'Brien", range: 'B5:D12' })
  assert.deepEqual(cellCitationTarget('Sheet1!XFD1048576'), { sheet: 'Sheet1', range: 'XFD1048576' })
  for (const address of ['A1', 'Sheet1!A:A', 'Sheet1!1:1', 'Sheet1!A0', 'Sheet1!XFE1', 'Sheet1!A1048577',
    'Sheet1!D12:B5', 'https://example.com', "'[Book.xlsm]Sheet1'!A1", "'bad/name'!A1", "'bad'quote'!A1",
    'Sheet1!A1,Sheet1!B2', 'Sheet1!A1\n', "'Modell Ü'!A1;alert(1)"]) {
    assert.throws(() => cellCitationTarget(address), { message: 'navigation_unavailable' })
  }
})

test('cell navigation activates and selects the exact validated range without writing cells', async () => {
  const calls = []
  await navigateCellCitation({ address: "'Modell Ü'!B5:D12", excel: { run: async action => action({
    workbook: { worksheets: { getItem: name => ({ activate: () => calls.push(['activate', name]),
      getRange: range => ({ select: () => calls.push(['select', range]) }) }) } },
    sync: async () => { calls.push(['sync']) },
  }) } })
  assert.deepEqual(calls, [['activate', 'Modell Ü'], ['select', 'B5:D12'], ['sync']])
})

test('navigation never uses a citation from another pane binding or a forged target', async () => {
  const calls = []
  const host = bridgeHost(async () => ({}), async target => { calls.push(target.address) })
  const nav = { ...request, type: 'dsh/excel-context/navigate', workbookId: 'pane-1', address: 'Sheet1!B5:D12' }
  await host.send({ ...nav, workbookId: 'different-pane' })
  await host.send({ ...nav, address: 'https://example.com' })
  assert.deepEqual(calls, [])
  await host.send(nav)
  assert.deepEqual(calls, ['Sheet1!B5:D12'])
  assert.deepEqual(host.replies.at(-1).data, { type: 'dsh/excel-context/navigation', version: 1,
    requestId: 'request-1', ok: true })
  host.dispose()
})

test('document location changes rotate binding and invalidate a historical citation', async () => {
  const targets = []
  const host = bridgeHost(async ({ workbookId }) => ({ workbook: { id: workbookId } }), async target => targets.push(target.address))
  await host.send(request)
  assert.equal(host.replies.at(-1).data.context.workbook.id, 'pane-1')
  host.office.context.document.url = 'file:///C:/Models/Other.xlsm'
  await host.send({ ...request, type: 'dsh/excel-context/navigate', workbookId: 'pane-1', address: 'Sheet1!A1' })
  assert.deepEqual(targets, [])
  assert.equal(host.replies.at(-1).data.error, 'context_unavailable')
  await host.send(request)
  assert.equal(host.replies.at(-1).data.context.workbook.id, 'pane-2')
  host.dispose()
})

test('a changed document cannot publish an earlier capture or select after the Office batch starts', async () => {
  let finish
  const host = bridgeHost(() => new Promise(resolve => { finish = resolve }))
  const pending = host.send(request)
  host.office.context.document.url = ''
  finish({ workbook: { id: 'pane-1' } })
  await pending
  assert.equal(host.replies.at(-1).data.error, 'context_unavailable')
  let touched = false
  await assert.rejects(navigateCellCitation({ address: 'Sheet1!A1', verify: () => false,
    excel: { run: async action => action({ workbook: { get worksheets() { touched = true; return {} } } }) } }),
  { message: 'navigation_unavailable' })
  assert.equal(touched, false)
  host.dispose()
})

test('cloud query-only identity changes invalidate binding without sending those parameters', async () => {
  const host = bridgeHost(async ({ workbookId, office }) => ({ workbook: {
    id: workbookId, path: workbookLocation(office.context.document.url),
  } }), async () => { throw new Error('Historical navigation must be rejected first') })
  host.office.context.document.url = 'https://example.com/Doc.aspx?sourcedoc=one&token=private'
  await host.send(request)
  const id = host.replies.at(-1).data.context.workbook.id
  assert.equal(host.replies.at(-1).data.context.workbook.path, 'https://example.com/Doc.aspx')
  host.office.context.document.url = 'https://example.com/Doc.aspx?sourcedoc=two&token=private'
  await host.send({ ...request, type: 'dsh/excel-context/navigate', workbookId: id, address: 'Sheet1!A1' })
  assert.equal(host.replies.at(-1).data.error, 'context_unavailable')
  assert.equal(JSON.stringify(host.replies).includes('private'), false)
  host.dispose()
})

test('disposal denies selection by a navigation batch that has not started yet', async () => {
  const gate = Promise.withResolvers()
  let allowed
  const host = bridgeHost(async () => ({}), async ({ verify }) => {
    await gate.promise
    allowed = verify()
  })
  const pending = host.send({ ...request, type: 'dsh/excel-context/navigate', workbookId: 'pane-1', address: 'Sheet1!A1' })
  host.dispose()
  gate.resolve()
  await pending
  assert.equal(allowed, false)
  assert.deepEqual(host.replies, [])
})
