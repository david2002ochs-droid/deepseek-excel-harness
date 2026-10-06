# 本地 Excel 任务窗格

[English](README.md) | 中文

此 Windows Excel Desktop 入口嵌入原版 Harness Web UI。现有 `dsh web` profile 负责协调、设置和会话。包装层初始化 Office.js，并通过回环 HTTPS 代理显示该 UI，包括其 HTTP 和 WebSocket 传输。代理保留浏览器端的 Host，并按照[原版 Web 部署指南](../docs/user/guide/public-deployments.zh.md)的要求，为后端 cookie 添加 Secure。其 manifest（元数据清单）仅授予 `Restricted` 权限；此入口不添加工作簿工具。其 32 像素 manifest 图标由[原版 Harness 桌面图标](../apps/desktop/resources/icon.png)缩放生成。

## 启动

使用 Node.js 24 或更新版本、仓库固定版本的 pnpm，以及已安装的 Excel Desktop。在检出目录的根目录中，准备原版 Harness 和独立的小型包装层：

```powershell
npx --yes pnpm@11.7.0 install --frozen-lockfile
npx --yes pnpm@11.7.0 run build
npm --prefix excel ci
npm --prefix excel run certs
```

Microsoft 开发证书命令将私钥存储在 Git 之外的 `%USERPROFILE%\.office-addin-dev-certs` 中，并安装 localhost 证书。如使用其他证书目录，请设置 `EXCEL_TLS_DIRECTORY`。凭据、证书、测试工作簿、日志和截图均须保存在检出目录之外。

在第一个终端中，使用检出目录之外的独立 Harness 主目录，使此次启动不读取其他项目的 profile 或设置：

```powershell
$env:DSH_HOME = Join-Path (Split-Path (Get-Location)) 'work\excel-harness-home'
npx --yes pnpm@11.7.0 dsh --profile web --no-open --port 3080 --public-url https://localhost:3443 --trusted-host localhost:3443
```

等待 Harness 输出启动 URL。在第二个终端的提示处粘贴完整的 `https://localhost:3443/?token=...` URL，然后启动 HTTPS 入口：

```powershell
$env:EXCEL_HARNESS_URL = Read-Host 'Paste the Harness HTTPS startup URL'
npm --prefix excel start
```

包装层要求使用此 localhost 根 URL，且仅包含一个非空 token 参数。它仅在 Host 为 `localhost:3443` 时提供携带凭据的任务窗格，并将 URL 传给嵌入的原版 UI，使 Harness 可以将启动 token 换成其签名会话 cookie。请保持 URL 私密，并将其保存在被跟踪文件之外；重启 Harness 后，更新环境变量并重启包装层。

在第三个终端中，将加载项旁加载到新的 Microsoft 测试工作簿：

```powershell
npm --prefix excel run sideload
```

旁加载命令使用 [Microsoft Office 加载项调试工具](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-debugging)。未指定 `--document` 时，它会将 Microsoft 的任务窗格工作簿模板复制到系统临时目录中的新文件，并打开该一次性测试工作簿。请勿通过 `--document` 传入现有工作簿。如任务窗格关闭，请打开 **Home > Add-ins > Developer Add-ins > DeepSeek Harness (local)**。根据需要加宽窗格，以显示原版 Web UI。

运行 `npm --prefix excel run validate`，使用 Microsoft 验证器检查 manifest。

## 停止与限制

运行 `npm --prefix excel run stop`，移除此 manifest 的开发注册，然后使用 Ctrl+C 停止两个服务器终端。关闭一次性测试工作簿，不要保存。移除证书是独立操作，可能影响其他本地 Office 开发；日常清理时不要卸载共享证书。

服务器启动和任务窗格可见并不意味着模型运行成功。此入口不配置模型路由。在任何 Excel agent（智能体）测试之前，必须显式配置所有者的 OpenRouter 路由：`meta/muse-spark-1.3-contributor`，提供方 `meta`，禁用 fallbacks，`data_collection=allow`，single/classic，控制器 low/high，side low，配置输出上限 8192，且不设置静态推理（reasoning）覆盖值。常规的付费 Excel 运行还要求服务器端具有 `TYPESAFE_API_KEY`；如不可用，只验证服务器和 UI。
