# 本地 Excel 任务窗格

[English](README.md) | 中文

此 Windows Excel Desktop 入口嵌入原版 Harness Web UI。现有 `dsh web` profile 负责协调、设置和会话。包装层初始化 Office.js，并通过回环 HTTPS 代理显示该 UI，包括其 HTTP 和 WebSocket 传输。代理保留浏览器端的 Host，并按照[原版 Web 部署指南](../docs/user/guide/public-deployments.zh.md)的要求，为后端 cookie 添加 Secure。其 manifest（元数据清单）仅授予 `Restricted` 权限；此入口不添加工作簿工具。其 32 像素 manifest 图标由[原版 Harness 桌面图标](../apps/desktop/resources/icon.png)缩放生成。

交付目标是完整的 Windows 安装包：安装过程负责前置依赖、受信任的 HTTPS 和加载项注册；用户直接从 Excel 打开 Harness，所需服务自动启动或连接。下述入口是过渡性的自动开发启动方式，并非该安装包。最终安装后的使用流程不得要求终端命令、复制 token 或开发者旁加载。

## 启动

使用 Node.js 24 或更新版本、仓库固定版本的 pnpm，以及已安装的 Excel Desktop。在检出目录的根目录中，准备原版 Harness 和独立的小型包装层：

```powershell
npx --yes pnpm@11.7.0 install --frozen-lockfile
npx --yes pnpm@11.7.0 run build
npm --prefix excel ci
npm --prefix excel run certs
```

Microsoft 开发证书命令将私钥存储在 Git 之外的 `%USERPROFILE%\.office-addin-dev-certs` 中，并安装 localhost 证书。如使用其他证书目录，请设置 `EXCEL_TLS_DIRECTORY`。凭据、证书、测试工作簿、日志和截图均须保存在检出目录之外。

完成准备后，一条开发命令即可启动原版 Harness web profile、私下获取其启动 URL、启动 HTTPS 包装层，并将加载项旁加载到一次性测试工作簿中：

```powershell
npm --prefix excel start
```

使用任务窗格期间，请保持该终端打开。无需复制 token，也无需分别打开服务器终端。启动器使用独立的 `DSH_HOME`：`%LOCALAPPDATA%\DeepSeekHarnessExcel\Harness`；如需其他目录，请将 `EXCEL_DSH_HOME` 设置为检出目录之外的路径。它忽略继承的 `DSH_HOME`，并从子进程环境中移除名称含凭据关键词的变量；请在该独立 Harness 主目录中配置服务器凭据。启动输出仅保留在内存中，不转发到终端，因此失败时只报告对应阶段，不会暴露 token。每个阶段的截止时间为两分钟。端口 3080 或 3443 已被占用时，启动器会报错，不会复用或停止占用端口的进程。

包装层要求使用 localhost 根 URL，且仅包含一个非空 token 参数。它仅在 Host 为 `localhost:3443` 时提供携带凭据的任务窗格，并将 URL 传给嵌入的原版 UI，使 Harness 可以将启动 token 换成其签名会话 cookie。

旁加载命令使用 [Microsoft Office 加载项调试工具](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-debugging)。未指定 `--document` 时，它会将 Microsoft 的任务窗格工作簿模板复制到系统临时目录中的新文件，并打开该一次性测试工作簿。请勿通过 `--document` 传入现有工作簿。如任务窗格关闭，请打开 **Home > Add-ins > Developer Add-ins > DeepSeek Harness (local)**。根据需要加宽窗格，以显示原版 Web UI。

如果已在运行的 Excel 提示加载项不再可用，其开发注册可能尚未刷新。仅关闭新生成的一次性测试工作簿，不要保存，然后按 Win+R 并运行 `EXCEL.EXE /x "<generated workbook path>"`，将占位内容替换为 Excel 文件信息中显示的一次性工作簿路径。此操作只会在独立的 Excel 进程中重新打开该测试工作簿；请保持原有工作簿打开。

运行 `npm --prefix excel test`，执行隔离的模拟子进程生命周期测试；测试不会启动 Harness、HTTPS 包装层或 Excel。这些测试不能作为原生任务窗格、身份验证、HTTP 或 WebSocket 验收通过的依据。运行 `npm --prefix excel run validate`，使用 Microsoft 验证器检查 manifest。

## 停止与限制

运行 `npm --prefix excel run stop`，移除此 manifest 的开发注册，然后在启动终端中按 Ctrl+C。取消、启动失败或服务器意外退出时，启动器会停止仍在运行的自有子进程，并等待其终止。在 Windows 上，启动器对每个仍存活的自有 PID 执行 `%SystemRoot%\System32\taskkill.exe /T` 并等待完成。如果该工具失败，启动器仍会终止自有的直接子进程并等待其退出，然后报告固定的清理错误；后代进程可能尚未清理完毕。它无法处理父进程退出前已脱离的后代进程，也无法在启动终端被强制终止后执行清理。Excel 工作簿由用户自行管理；启动器不会关闭已有的 Excel 实例。关闭一次性测试工作簿，不要保存。移除证书是独立操作，可能影响其他本地 Office 开发；日常清理时不要卸载共享证书。

服务器启动和任务窗格可见并不意味着模型运行成功。此入口不配置模型路由。在任何 Excel agent（智能体）测试之前，必须显式配置所有者的 OpenRouter 路由：`meta/muse-spark-1.3-contributor`，提供方 `meta`，禁用 fallbacks，`data_collection=allow`，single/classic，控制器 low/high，side low，配置输出上限 8192，且不设置静态推理（reasoning）覆盖值。常规的付费 Excel 运行还要求服务器端具有 `TYPESAFE_API_KEY`；如不可用，只验证服务器和 UI。
