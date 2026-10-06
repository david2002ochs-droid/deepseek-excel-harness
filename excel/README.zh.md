# 本地 Excel 任务窗格

[English](README.md) | 中文

此 Windows Excel Desktop 入口嵌入原版 Harness Web UI。现有 `dsh web` profile 负责协调、设置和会话。包装层初始化 Office.js，并通过回环 HTTPS 代理显示该 UI，包括其 HTTP 和 WebSocket 传输。代理保留浏览器端的 Host，并按照[原版 Web 部署指南](../docs/user/guide/public-deployments.zh.md)的要求，为后端 cookie 添加 Secure。manifest 更新会提高版本号，以便 Office 刷新缓存的权限。其 manifest（元数据清单）请求 Microsoft 的 Excel 专用 API 所要求的 `ReadWriteDocument` 权限；包装层仅读取工作簿与选择范围的元数据，不添加工作簿执行工具。其 32 像素 manifest 图标由[原版 Harness 桌面图标](../apps/desktop/resources/icon.png)缩放生成。

交付目标是完整的 Windows 安装包：安装过程负责前置依赖、受信任的 HTTPS 和带版本的加载项注册更新；用户直接从 Excel 打开 Harness，所需服务自动启动或连接。下述入口是过渡性的自动开发启动方式，并非该安装包。最终安装后的使用流程不得要求终端命令、复制 token、开发者旁加载或用户清理缓存；验收包括原生工作簿上下文捕获。

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

旁加载辅助程序使用 [Microsoft Office 加载项开发设置](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-dev-settings)的公开 API。它注册当前 manifest，禁用开发调试和实时重载，在独立临时目录中生成 Microsoft 的任务窗格工作簿模板，并在打开前为工作簿设置唯一文件名。重复启动会保留先前的测试工作簿，并保持现有 Excel 实例打开。生成的文件在打开后继续保留，打开失败时也保留。如开发过程中任务窗格关闭，请打开 **Home > Add-ins > Developer Add-ins > DeepSeek Harness (local)**。根据需要加宽窗格，以显示原版 Web UI。

Office 可能保留已打开任务窗格的 manifest 权限。修改注册或权限时，提高 manifest 版本并测试新的临时工作簿；关闭先前的测试工作簿时保留其修改。现有用户工作簿保持打开。注册和权限刷新由安装后的更新过程负责。

运行 `npm --prefix excel test`，执行隔离的模拟子进程生命周期、只读元数据桥接，以及替换注册和打开操作的真实 Office 工作簿生成测试；测试不会启动 Harness、HTTPS 包装层或 Excel。这些测试不能作为原生任务窗格、身份验证、HTTP 或 WebSocket 验收通过的依据。运行 `npm --prefix excel run validate`，使用 Microsoft 验证器检查 manifest。

## 工作簿上下文

Excel 启动通过 `context.patch.yml`，将 Standard 固定为唯一内置代理预设和默认值，禁用预设选择界面，同时让现有会话插件显式启用父框架桥接。Minimal、PTC 和 Creator 的声明被禁用；选择了这些预设的会话需要重新启用相应声明，才能在重启后恢复。Host 在浏览器插件启动前注入公开的 origin 与超时；该启动值不含凭据或工作簿数据。部署配置变更后需重新加载任务窗格。每条普通提交的消息在异步附件准备之前请求新的 Office 元数据读取。第一条消息在首个模型请求之前提供工作簿定位信息；后续消息携带各自捕获的选择范围。捕获内容包括窗格绑定的工作簿标识符、通过 ExcelApi 1.7 获取的工作簿名称（旧主机使用文件定位信息的末段），以及可用时不含凭据的文件定位信息、活动工作表标识符及名称、带工作表名称的选择地址（受支持主机上的多个区域），以及观察时间。JSON 最多为 2048 个 UTF-8 字节。不收集单元格值、公式、VBA 源码、完整工作表清单或工具目录。

元数据保留在普通用户消息日志及其模型投影中。聊天和排队消息将捕获的选择范围显示为小型引用标签。精确的来源、父子窗口校验及每次提交的请求标识符关联桥接通信。元数据不可用、超时或过大时，Excel 提交被阻止且草稿保留；选择目标不会被截断。独立 Web 会话不依赖此桥接。

助手正文保留原有的单元格引用语法，例如 `[[cite:'Revenue 2027'!B4:D8]]`。在绑定的 Excel 窗格中，有效引用会选择对应工作表和区域。导航验证原消息的窗格标识符及带工作表名称、有界的 A1 地址，拒绝外部工作簿引用和可执行文本。无效标记、代码示例以及没有匹配工作簿上下文的引用保持为普通文本。点击引用会改变活动工作表及选择范围，但不会写入单元格。引用记号本身不证明智能体读取或修改了这些单元格。

标识符仅属于当前任务窗格绑定，不是原生 COM 工作簿句柄。重新加载窗格或 Office 报告不同文档位置时（包括另存为），标识符会改变；导航在选择之前再次检查位置。Office 可能无法为尚未保存的工作簿提供文件定位信息，此时绑定依赖 Office 将窗格保持在原文档中。第二个重叠的捕获或导航立即失败并保留草稿，不读取后来的选择。此上下文既不建立 COM 连接，也不安装 Excel/VBA CLI；后续 COM 操作必须解析并验证捕获的目标，包括观察后发生的变化。工作簿身份与选择范围的概念参考现有 `vba-excel-cre-auditor` 集成，但不导入其大型准备扫描或运行时。

## 停止与限制

运行 `npm --prefix excel run stop`，移除此 manifest 的开发注册，然后在启动终端中按 Ctrl+C。取消、启动失败或服务器意外退出时，启动器会停止仍在运行的自有子进程，并等待其终止。在 Windows 上，启动器对每个仍存活的自有 PID 执行 `%SystemRoot%\System32\taskkill.exe /T` 并等待完成。如果该工具失败，启动器仍会终止自有的直接子进程并等待其退出，然后报告固定的清理错误；后代进程可能尚未清理完毕。它无法处理父进程退出前已脱离的后代进程，也无法在启动终端被强制终止后执行清理。Excel 工作簿由用户自行管理；启动器不会关闭已有的 Excel 实例。关闭一次性测试工作簿，不要保存。移除证书是独立操作，可能影响其他本地 Office 开发；日常清理时不要卸载共享证书。

服务器启动和任务窗格可见并不意味着模型运行成功。此入口不配置模型路由。在任何 Excel agent（智能体）测试之前，必须显式配置所有者的 OpenRouter 路由：`meta/muse-spark-1.3-contributor`，提供方 `meta`，禁用 fallbacks，`data_collection=allow`，single/classic，控制器 low/high，side low，配置输出上限 8192，且不设置静态推理（reasoning）覆盖值。常规的付费 Excel 运行还要求服务器端具有 `TYPESAFE_API_KEY`；如不可用，只验证服务器和 UI。
