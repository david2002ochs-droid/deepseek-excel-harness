/** Loopback HTTPS entry for the Office metadata wrapper and the dsh Web profile. */
import { readFileSync } from 'node:fs'
import { createServer } from 'node:https'
import { homedir } from 'node:os'
import { join } from 'node:path'
import httpProxy from 'http-proxy'

const certificateDirectory = process.env.EXCEL_TLS_DIRECTORY ?? join(homedir(), '.office-addin-dev-certs')
let harnessUrl
try {
  harnessUrl = new URL(process.env.EXCEL_HARNESS_URL ?? '')
} catch (_error) {
  throw new Error('Set EXCEL_HARNESS_URL to the HTTPS startup URL printed by dsh Web.')
}
if (harnessUrl.origin !== 'https://localhost:3443' || harnessUrl.pathname !== '/'
  || harnessUrl.username || harnessUrl.password || harnessUrl.hash
  || harnessUrl.searchParams.size !== 1 || !harnessUrl.searchParams.get('token')) {
  throw new Error('EXCEL_HARNESS_URL must be https://localhost:3443/?token=<launch token>.')
}
const taskpane = readFileSync(new URL('./taskpane.html', import.meta.url), 'utf8')
  .replace('__EXCEL_HARNESS_URL__', JSON.stringify(harnessUrl.href).replaceAll('<', '\\u003c'))
const icon = readFileSync(new URL('./icon.png', import.meta.url))
const contextModule = readFileSync(new URL('./context.mjs', import.meta.url))
const proxy = httpProxy.createProxyServer({ target: 'http://127.0.0.1:3080', ws: true, changeOrigin: false })

proxy.on('proxyRes', (proxyResponse) => {
  const cookies = proxyResponse.headers['set-cookie']
  if (cookies) {
    proxyResponse.headers['set-cookie'] = cookies.map(cookie => `${cookie}; Secure`)
  }
})

proxy.on('error', (_error, _request, response) => {
  if ('writeHead' in response && !response.headersSent && !response.destroyed) {
    response.writeHead(502, { 'content-type': 'text/plain; charset=utf-8' })
    response.end('Start the local dsh Web profile on port 3080.')
  } else {
    response.destroy()
  }
})

const server = createServer({
  key: readFileSync(join(certificateDirectory, 'localhost.key')),
  cert: readFileSync(join(certificateDirectory, 'localhost.crt')),
}, (request, response) => {
  const pathname = new URL(request.url, 'https://localhost:3443').pathname
  if (pathname === '/excel/taskpane.html') {
    if (request.headers.host !== 'localhost:3443') {
      response.writeHead(403, { 'content-type': 'text/plain; charset=utf-8' })
      response.end('Open the Excel entry at https://localhost:3443.')
      return
    }
    response.writeHead(200, { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store' })
    response.end(taskpane)
    return
  }
  if (pathname === '/excel/icon.png') {
    response.writeHead(200, { 'content-type': 'image/png' })
    response.end(icon)
    return
  }
  if (pathname === '/excel/context.mjs') {
    response.writeHead(200, { 'content-type': 'text/javascript; charset=utf-8', 'cache-control': 'no-store' })
    response.end(contextModule)
    return
  }
  proxy.web(request, response)
})

server.on('upgrade', (request, socket, head) => proxy.ws(request, socket, head))
server.listen(3443, '127.0.0.1', () => {
  console.log('Excel taskpane: https://localhost:3443/excel/taskpane.html')
})
