/** Office metadata and cell-selection bridge for the embedded Harness composer. */
export const EXCEL_CONTEXT_BYTES = 2048

/** Validate a single sheet-qualified bounded A1 target without accepting executable text. */
export function cellCitationTarget(address) {
  if (typeof address !== 'string' || address !== address.trim() || address.length > 256) {
    throw new Error('navigation_unavailable')
  }
  const match = /^(?:'((?:[^'\r\n]|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!(\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})(?::(\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}))?$/u.exec(address)
  if (!match) throw new Error('navigation_unavailable')
  const sheet = match[1] ? match[1].replaceAll("''", "'") : match[2]
  if (!sheet || sheet.length > 31 || /^[']|[']$/u.test(sheet) || /[\\/:?*\[\]\u0000-\u001f]/u.test(sheet)) {
    throw new Error('navigation_unavailable')
  }
  const cell = value => {
    const parts = /^([A-Za-z]+)([0-9]+)$/u.exec(value.replaceAll('$', ''))
    const letters = parts[1].toUpperCase()
    const column = [...letters].reduce((result, letter) => result * 26 + letter.charCodeAt(0) - 64, 0)
    const row = Number(parts[2])
    if (column > 16384 || row > 1048576) throw new Error('navigation_unavailable')
    return { text: `${letters}${row}`, column, row }
  }
  const first = cell(match[3])
  const last = cell(match[4] ?? match[3])
  if (first.column > last.column || first.row > last.row) throw new Error('navigation_unavailable')
  return { sheet, range: match[4] ? `${first.text}:${last.text}` : first.text }
}

/** Navigate within the bound Office workbook without writing cells or evaluating code. */
export async function navigateCellCitation({ excel, address, verify = () => true }) {
  const target = cellCitationTarget(address)
  await excel.run(async context => {
    if (!verify()) throw new Error('navigation_unavailable')
    const worksheet = context.workbook.worksheets.getItem(target.sheet)
    worksheet.activate()
    worksheet.getRange(target.range).select()
    await context.sync()
  })
}

/** Convert Office's document location into a credential-free workbook locator. */
export function workbookLocation(value) {
  if (typeof value !== 'string' || !value.trim()) return undefined
  const location = value.trim()
  if (/^(?:[a-z]:[\\/]|\\\\)/iu.test(location)) return location
  try {
    const url = new URL(location)
    if (url.protocol === 'file:') {
      const path = decodeURIComponent(url.pathname).replaceAll('/', '\\')
      return url.hostname ? `\\\\${url.hostname}${path}` : path.replace(/^\\(?=[a-z]:)/iu, '')
    }
    if (url.protocol !== 'https:' && url.protocol !== 'http:') return undefined
    url.username = ''
    url.password = ''
    url.search = ''
    url.hash = ''
    return url.href
  } catch (_error) {
    // An unavailable or malformed Office URL is not a file-system identity.
    return undefined
  }
}

/** Capture only host-owned workbook, worksheet and selection metadata in one batch. */
export async function captureExcelContext({ office, excel, workbookId, now = () => new Date() }) {
  const path = workbookLocation(office.context.document.url)
  return excel.run(async context => {
    const supportsName = office.context.requirements.isSetSupported('ExcelApi', '1.7')
    if (supportsName) context.workbook.load('name')
    const worksheet = context.workbook.worksheets.getActiveWorksheet()
    worksheet.load('id,name')
    const selection = office.context.requirements.isSetSupported('ExcelApi', '1.9')
      ? context.workbook.getSelectedRanges()
      : context.workbook.getSelectedRange()
    selection.load('address')
    await context.sync()
    if (typeof worksheet.id !== 'string' || !worksheet.id
      || typeof worksheet.name !== 'string' || !worksheet.name
      || typeof selection.address !== 'string' || !selection.address) {
      throw new Error('context_unavailable')
    }
    const name = supportsName ? context.workbook.name : path?.split(/[\\/]/u).at(-1)
    if (supportsName && (typeof name !== 'string' || !name)) throw new Error('context_unavailable')
    const captured = {
      workbook: { id: workbookId, ...(name ? { name } : {}), ...(path ? { path } : {}) },
      worksheet: { id: worksheet.id, name: worksheet.name },
      selection: { workbookId, address: selection.address },
      observedAt: now().toISOString(),
    }
    if (new TextEncoder().encode(JSON.stringify(captured)).byteLength > EXCEL_CONTEXT_BYTES) {
      throw new Error('context_too_large')
    }
    return captured
  })
}

/** Bind capture requests to this frame and its exact HTTPS origin; return a disposer. */
export function installExcelContextBridge({ window, frame, office, excel, origin,
  workbookId = window.crypto.randomUUID(), capture = captureExcelContext, navigate = navigateCellCitation }) {
  let disposed = false
  let reading = false
  // Compare the host location privately: cloud document IDs can live in query
  // parameters that must never enter the credential-free model locator.
  const documentLocation = () => office.context.document.url ?? ''
  let boundLocation = documentLocation()
  let boundId = workbookId
  const currentBinding = () => {
    const location = documentLocation()
    if (location !== boundLocation) {
      boundLocation = location
      boundId = window.crypto.randomUUID()
    }
    return boundId
  }
  const onMessage = async event => {
    if (disposed || event.source !== frame.contentWindow || event.origin !== origin) return
    const request = event.data
    if (!request || typeof request !== 'object'
      || !['dsh/excel-context/request', 'dsh/excel-context/navigate'].includes(request.type)
      || request.version !== 1 || typeof request.requestId !== 'string'
      || !/^[a-z0-9_-]{1,128}$/iu.test(request.requestId)) return
    const navigation = request.type === 'dsh/excel-context/navigate'
    const reply = data => {
      if (!disposed) frame.contentWindow.postMessage({ type: navigation
        ? 'dsh/excel-context/navigation' : 'dsh/excel-context/response',
        version: 1, requestId: request.requestId, ...data }, origin)
    }
    if (reading) { reply({ error: 'context_unavailable' }); return }
    reading = true
    try {
      const id = currentBinding()
      if (navigation) {
        if (request.workbookId !== id) throw new Error('navigation_unavailable')
        cellCitationTarget(request.address)
        await navigate({ excel, address: request.address, verify: () => !disposed && currentBinding() === id })
        reply({ ok: true })
      } else {
        const context = await capture({ office, excel, workbookId: id })
        if (currentBinding() !== id) throw new Error('context_unavailable')
        reply({ context })
      }
    } catch (error) {
      reply({ error: error instanceof Error && error.message === 'context_too_large'
        ? 'context_too_large' : 'context_unavailable' })
    } finally {
      reading = false
    }
  }
  window.addEventListener('message', onMessage)
  return () => {
    disposed = true
    window.removeEventListener('message', onMessage)
  }
}
