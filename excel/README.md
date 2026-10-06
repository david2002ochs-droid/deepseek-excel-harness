# Local Excel taskpane

English | [中文](README.zh.md)

This Windows Excel Desktop entry embeds the original Harness Web UI. The existing `dsh web` profile owns orchestration, settings, and sessions. The wrapper initializes Office.js and displays that UI through a loopback HTTPS proxy, including its HTTP and WebSocket transport. The proxy preserves the browser-facing Host and adds Secure to backend cookies as required by the [original Web deployment guide](../docs/user/guide/public-deployments.md). Its manifest grants only `Restricted` permissions; this entry adds no workbook tools. Its 32-pixel manifest icon is resized from the [original Harness desktop icon](../apps/desktop/resources/icon.png).

The delivery target is a complete Windows installation package: installation owns prerequisites, trusted HTTPS, and add-in registration, and users open Harness directly from Excel while services start or connect automatically. The entry below is an interim automatic development launch, not that installer. The final installed experience must require no terminal commands, token copying, or developer sideloading.

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

The sideload command uses [Microsoft's Office add-in debugging tooling](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-debugging). With no `--document`, it copies Microsoft's taskpane workbook template to a new file in the system temporary directory and opens that disposable workbook. Do not pass an existing workbook with `--document`. If the taskpane is closed, open **Home > Add-ins > Developer Add-ins > DeepSeek Harness (local)**. Widen the pane as needed for the original Web UI.

If an already-running Excel reports that the add-in is no longer available, its developer registration may be stale. Close only the newly generated disposable test workbook without saving, then press Win+R and run `EXCEL.EXE /x "<generated workbook path>"`, replacing the placeholder with the disposable workbook path shown in Excel’s file information. This reopens only that test workbook in a separate Excel process; leave your original workbooks open.

Run `npm --prefix excel test` for isolated fake-child lifecycle tests; they do not launch Harness, the HTTPS wrapper, or Excel. These tests do not establish native taskpane, authentication, HTTP, or WebSocket acceptance. Run `npm --prefix excel run validate` to check the manifest with Microsoft's validator.

## Stop and limits

Run `npm --prefix excel run stop` to remove this manifest's development registration, then press Ctrl+C in the launch terminal. Cancellation, startup failure, and unexpected server exit stop live owned child processes and await their termination. On Windows, the launcher invokes `%SystemRoot%\System32\taskkill.exe /T` for each live owned PID and awaits completion. If that utility fails, it still terminates and awaits the owned direct child before reporting a fixed cleanup error; descendant cleanup may remain incomplete. It cannot recover descendants orphaned before their parent exits, or clean up after the launch terminal is forcibly killed. Keep Excel workbooks under your control; the launcher does not close existing Excel instances. Close the disposable test workbook without saving. Certificate removal is separate and may affect other local Office development; do not uninstall shared certificates as routine cleanup.

Server startup and a visible taskpane do not establish a successful model run. This entry does not configure a model route. Before any Excel agent test, explicitly configure the owner's OpenRouter route: `meta/muse-spark-1.3-contributor`, provider `meta`, fallbacks disabled, `data_collection=allow`, single/classic, controller low/high, side low, configured output cap 8192, and no static reasoning override. A regular paid Excel run also requires the server-side `TYPESAFE_API_KEY`; if unavailable, restrict verification to server and UI checks.
