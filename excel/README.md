# Local Excel taskpane

English | [中文](README.zh.md)

This Windows Excel Desktop entry embeds the original Harness Web UI. The existing `dsh web` profile owns orchestration, settings, and sessions. The wrapper initializes Office.js and displays that UI through a loopback HTTPS proxy, including its HTTP and WebSocket transport. The proxy preserves the browser-facing Host and adds Secure to backend cookies as required by the [original Web deployment guide](../docs/user/guide/public-deployments.md). Its manifest grants only `Restricted` permissions; this entry adds no workbook tools. Its 32-pixel manifest icon is resized from the [original Harness desktop icon](../apps/desktop/resources/icon.png).

## Start

Use Node.js 24 or newer, the repository's pinned pnpm, and an installed Excel Desktop. From the checkout root, prepare the original Harness and the small independent wrapper:

```powershell
npx --yes pnpm@11.7.0 install --frozen-lockfile
npx --yes pnpm@11.7.0 run build
npm --prefix excel ci
npm --prefix excel run certs
```

The Microsoft development-certificate command stores its private key outside Git in `%USERPROFILE%\.office-addin-dev-certs` and installs the localhost certificate. Set `EXCEL_TLS_DIRECTORY` if using a different certificate directory. Keep credentials, certificates, test workbooks, logs, and screenshots outside the checkout.

In a first terminal, use a separate Harness home outside the checkout so this launch does not read another project's profiles or settings:

```powershell
$env:DSH_HOME = Join-Path (Split-Path (Get-Location)) 'work\excel-harness-home'
npx --yes pnpm@11.7.0 dsh --profile web --no-open --port 3080 --public-url https://localhost:3443 --trusted-host localhost:3443
```

Wait for Harness to print its startup URL. In a second terminal, paste that complete `https://localhost:3443/?token=...` URL at the prompt, then start the HTTPS entry:

```powershell
$env:EXCEL_HARNESS_URL = Read-Host 'Paste the Harness HTTPS startup URL'
npm --prefix excel start
```

The wrapper requires that exact localhost root URL with one nonempty token parameter. It serves the credential-bearing taskpane only for Host `localhost:3443` and supplies the URL to the embedded original UI so Harness can exchange its launch token for its signed session cookie. Keep the URL private and outside tracked files; after restarting Harness, update the environment variable and restart the wrapper.

In a third terminal, sideload into a fresh Microsoft test workbook:

```powershell
npm --prefix excel run sideload
```

The sideload command uses [Microsoft's Office add-in debugging tooling](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-debugging). With no `--document`, it copies Microsoft's taskpane workbook template to a new file in the system temporary directory and opens that disposable workbook. Do not pass an existing workbook with `--document`. If the taskpane is closed, open **Home > Add-ins > Developer Add-ins > DeepSeek Harness (local)**. Widen the pane as needed for the original Web UI.

Run `npm --prefix excel run validate` to check the manifest with Microsoft's validator.

## Stop and limits

Run `npm --prefix excel run stop` to remove this manifest's development registration, then stop the two server terminals with Ctrl+C. Close the disposable test workbook without saving. Certificate removal is separate and may affect other local Office development; do not uninstall shared certificates as routine cleanup.

Server startup and a visible taskpane do not establish a successful model run. This entry does not configure a model route. Before any Excel agent test, explicitly configure the owner's OpenRouter route: `meta/muse-spark-1.3-contributor`, provider `meta`, fallbacks disabled, `data_collection=allow`, single/classic, controller low/high, side low, configured output cap 8192, and no static reasoning override. A regular paid Excel run also requires the server-side `TYPESAFE_API_KEY`; if unavailable, restrict verification to server and UI checks.
