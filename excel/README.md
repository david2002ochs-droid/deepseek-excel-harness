# Local Excel taskpane

English | [中文](README.zh.md)

This Windows Excel Desktop entry embeds the original Harness Web UI. The existing `dsh web` profile owns orchestration, settings, and sessions. The wrapper initializes Office.js and displays that UI through a loopback HTTPS proxy, including its HTTP and WebSocket transport. The proxy preserves the browser-facing Host and adds Secure to backend cookies as required by the [original Web deployment guide](../docs/user/guide/public-deployments.md). Manifest updates advance the version so Office refreshes cached permissions. Its manifest requests `ReadWriteDocument`, which Microsoft requires for application-specific Excel APIs; the wrapper only reads workbook and selection metadata and adds no workbook execution tools. Its 32-pixel manifest icon is resized from the [original Harness desktop icon](../apps/desktop/resources/icon.png).

The delivery target is a complete Windows installation package: installation owns prerequisites, trusted HTTPS, and versioned add-in registration updates, and users open Harness directly from Excel while services start or connect automatically. The entry below is an interim automatic development launch, not that installer. The final installed experience must require no terminal commands, token copying, developer sideloading, or user cache cleanup; acceptance includes native workbook-context capture.

## Start

Use Node.js 24 or newer, the repository's pinned pnpm, and an installed Excel Desktop. From the checkout root, prepare the original Harness and the small independent wrapper:

```powershell
npx --yes pnpm@11.7.0 install --frozen-lockfile
npx --yes pnpm@11.7.0 run build
npm --prefix excel ci
npm --prefix excel run certs
```

The Microsoft development-certificate command stores its private key outside Git in `%USERPROFILE%\.office-addin-dev-certs` and installs the localhost certificate. Set `EXCEL_TLS_DIRECTORY` if using a different certificate directory. Keep credentials, certificates, test workbooks, logs, and screenshots outside the checkout.

After setup, one development command starts the original Harness web profile, captures its startup URL privately, starts the HTTPS wrapper, and sideloads a disposable test workbook:

```powershell
npm --prefix excel start
```

Keep that terminal open while using the pane. No token copying or separate server terminals are required. The launcher uses a dedicated `DSH_HOME` at `%LOCALAPPDATA%\DeepSeekHarnessExcel\Harness`; set `EXCEL_DSH_HOME` to another directory outside the checkout if needed. It ignores an inherited `DSH_HOME` and strips credential-named environment variables from children; configure server credentials in that dedicated Harness home. Startup output is retained only in memory and never forwarded, so failures report the affected stage without exposing tokens. Each stage has a two-minute deadline. Occupied ports 3080 or 3443 cause a failure; the launcher does not reuse or stop their owners.

The wrapper requires the localhost root URL with one nonempty token parameter. It serves the credential-bearing taskpane only for Host `localhost:3443` and supplies the URL to the embedded original UI so Harness can exchange its launch token for its signed session cookie.

The sideload helper uses the public APIs of [Microsoft's Office add-in development settings](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-dev-settings). It registers the current manifest, disables development debugging and live reload, generates Microsoft's taskpane workbook template in an isolated temporary directory, and gives the workbook a unique basename before opening it. Repeated launches preserve earlier test workbooks and leave existing Excel instances open. Generated files remain available after opening, including when opening fails. If the taskpane is closed during development, open **Home > Add-ins > Developer Add-ins > DeepSeek Harness (local)**. Widen the pane as needed for the original Web UI.

Office can retain the manifest permissions of an already-open taskpane. Advance the manifest version when changing registration or permissions and test a fresh disposable workbook; preserve any changes when closing an earlier test workbook. Existing user workbooks remain open. Registration and permission refresh belong to the installed update process.

Run `npm --prefix excel test` for isolated fake-child lifecycle, read-only metadata bridge, and real Office workbook-generation tests with registration and opening replaced; they do not launch Harness, the HTTPS wrapper, or Excel. These tests do not establish native taskpane, authentication, HTTP, or WebSocket acceptance. Run `npm --prefix excel run validate` to check the manifest with Microsoft's validator.

## Workbook context

The launcher sets Standard's persona prefix to `You are an Excel agent.` through `DSH_EXCEL_PERSONA_PREFIX`. Tool guidance and skills remain available; this prefix does not replace the complete system prompt. User profile overrides of Standard's persona remain effective.

The Excel launch applies `context.patch.yml` to fix Standard as the only built-in agent preset and default, disable the preset selection UI, and opt the existing conversation plugin into the parent-frame bridge. Minimal, PTC and Creator declarations are disabled; sessions that selected those presets require their declarations to be enabled again before resuming after restart. The Host injects the public origin and timeout into the browser before plugins start; this boot value contains no credentials or workbook data. Reload the taskpane after deployment configuration changes. Each ordinary submitted message requests a fresh Office metadata read before asynchronous attachment preparation. Its first message gives the agent the workbook orientation before the first model request; later messages carry their own captured selection. The capture includes a pane-bound workbook identifier, the workbook name from ExcelApi 1.7 (a file-locator basename on older hosts), a credential-free file locator when available, active worksheet identity/name, the sheet-qualified selected address (including multiple areas on supported hosts), and observation time. JSON is limited to 2048 UTF-8 bytes. No cell values, formulas, VBA source, complete sheet inventory or tool catalog are collected.

The metadata stays in the ordinary logged user message and its model projection. Chat and queued messages display the captured selection as a small reference chip. Exact origin, parent/child window checks and per-submission request identifiers correlate bridge traffic. Unavailable, timed-out or oversized metadata blocks that Excel submission and preserves the draft; selected targets are never truncated. Ordinary standalone Web conversation remains independent of the bridge.

Assistant prose retains the original cell-citation syntax, such as `[[cite:'Revenue 2027'!B4:D8]]`. In the bound Excel pane, a valid citation selects that worksheet and range. Navigation verifies the originating message's pane identifier and a bounded, sheet-qualified A1 address; external workbook references and executable text are rejected. Invalid tokens, code examples and citations without a matching workbook context remain ordinary text. Selecting a citation changes the active worksheet and selection, but does not write cells. Citation notation is a reference, not proof that the agent read or changed those cells.

The identifier belongs to the current taskpane binding; it is not a native COM workbook handle and changes when the pane reloads or Office reports a different document location, including Save As. Navigation rechecks that location before selecting. Office may provide no file locator for an unsaved workbook; such a binding relies on Office keeping the pane attached to that document. A second overlapping capture or navigation fails immediately and preserves its draft instead of reading a later selection. The context does not establish a COM connection or install an Excel/VBA CLI: a later COM operation must resolve and verify the captured target, including any changes since observation. Workbook identity and selection concepts follow the existing `vba-excel-cre-auditor` integration, without importing its large preparation scan or runtime.

## Stop and limits

Run `npm --prefix excel run stop` to remove this manifest's development registration, then press Ctrl+C in the launch terminal. Cancellation, startup failure, and unexpected server exit stop live owned child processes and await their termination. On Windows, the launcher invokes `%SystemRoot%\System32\taskkill.exe /T` for each live owned PID and awaits completion. If that utility fails, it still terminates and awaits the owned direct child before reporting a fixed cleanup error; descendant cleanup may remain incomplete. It cannot recover descendants orphaned before their parent exits, or clean up after the launch terminal is forcibly killed. Keep Excel workbooks under your control; the launcher does not close existing Excel instances. Close the disposable test workbook without saving. Certificate removal is separate and may affect other local Office development; do not uninstall shared certificates as routine cleanup.

Server startup and a visible taskpane do not establish a successful model run. This entry does not configure a model route. Before any Excel agent test, explicitly configure the owner's OpenRouter route: `meta/muse-spark-1.3-contributor`, provider `meta`, fallbacks disabled, `data_collection=allow`, single/classic, controller low/high, side low, configured output cap 8192, and no static reasoning override. A regular paid Excel run also requires the server-side `TYPESAFE_API_KEY`; if unavailable, restrict verification to server and UI checks.
