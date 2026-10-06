---
description: 为 Excel 任务窗格 VBA skill 提供受保护的实时工作簿绑定和私有 xlflow 源码项目。
kind: package-library
---

# Excel VBA 辅助程序参考

[English](README.md) | 中文

## 摘要

[skill](SKILL.md) 使用 [workbook.mjs](scripts/workbook.mjs) 验证任务窗格捕获的本地定位信息，并操作唯一打开的原始工作簿。[inventory.ps1](scripts/inventory.ps1) 在每个 Excel 进程中通过原生 Windows 句柄和 COM 读取工作簿身份。清单读取和验证均不会打开、关闭、保存工作簿，也不会修改工作簿或安全设置。产品将整个 skill 目录复制到 `<home>/skills/excel-vba`；辅助程序使用 `<home>/runtime/xlflow-0.35.0`，并在执行前检查 CLI（命令行界面）和桥接程序的哈希值。

## 请求与源码归属

运行 `node "<installed skill>/scripts/workbook.mjs" "<request.json>"`。请求包含 `action`、当前请求的任务窗格 `context`，以及用于命令的字符串数组 `args`。上下文使用 `{workbook:{id,path},worksheet:{id,name},selection:{workbookId,address},observedAt}`。成功结果返回 `ok`、目标身份，以及适用时的 `project`、`exitCode` 和完整的 xlflow JSON `output`。拒绝结果返回 `ok:false`、`error.code` 和说明修正方法的 `error.message`；进程退出码非零。模块导出 `validateBinding(context, operations)` 和 `workbookRequest(request, operations)` 供集成和测试使用；生产调用方可以提供 `operations.home`，其余操作应保持原生实现。

| 操作 | 行为 |
| --- | --- |
| `validate` | 检查本地文件和格式，读取完整实时清单；不写入文件，也不调用 xlflow。 |
| `prepare` | 原子创建私有项目或检查已有的托管项目；不附加或拉取。 |
| `bootstrap` | 创建新项目，在附加前重新验证，然后在通过 Excel/session 拉取前再次验证。绝不覆盖已有项目。 |
| `rebind` | 验证后显式更新辅助程序管理的任务窗格绑定；保留源码、配置和会话，拒绝冲突会话。 |
| `command` | 调用支持的源码或实时命令前重新验证。源码 `inspect calls/symbols` 不检查会话，也不传递会话参数；工作簿观察要求外部会话。参数作为独立进程参数传递；拒绝覆盖目标、使用文件后端或管理生命周期。 |

项目位于 `<home>/vba-projects/<path-hash>`，包含精确的 `[excel].path`、真实的 `src/{modules,classes,forms,workbook}` 目录、UserForm sidecar 模式以及辅助程序管理的 `binding.json`。创建过程使用同级暂存目录和原子重命名。复用要求配置仍由辅助程序管理，且工作簿和进程相同；自定义或无关配置会导致拒绝，不会被替换。同一工作簿的辅助程序调用必须串行执行。需要会话的命令必须匹配 `.xlflow/session.json` 的路径、PID 和 external 所有者。已有会话继续使用，绝不静默重新附加。拉取只接受空源码目录，因此后续源码协调需要在自动初始化之外显式决定权威来源。

保护逻辑在读取清单前拒绝分开传递或使用等号形式的 `--input` 和 `--save-as`。支持的命令中，只有 `export-image` 接受工作簿位置参数；其文档规定的选项参数数量用于区分选项值与被拒绝的位置参数。不会根据文件名后缀拒绝宏参数或源码路径。`run --save` 仍表示明确请求在执行成功后保存原始工作簿。

## 限制与验证

仅接受本地驱动器上的 `.xlsm`、`.xlsb`、`.xlam` 和 `.xltm` 定位信息，且实时 `FileFormat` 必须匹配。已保存原始工作簿中的未保存编辑仍然有效。缺失、云端、网络或未保存名称定位信息、重复匹配、无法访问的 Excel 进程，以及变化的进程清单都会导致拒绝。必须能够访问 Windows 原生对象模型；不能假定不可见或无法访问的实例为空。捕获的定位信息无法在另存为后确定当前任务窗格身份，即使旧路径仍匹配另一打开的工作簿。身份不确定时请求新捕获的任务窗格元数据；TTL 不能替代身份验证。清单读取与 CLI 执行之间仍可能发生用户操作竞争。

辅助程序将本地清单脚本作为静态编码的 PowerShell 命令传递，不修改执行策略。现有 shell 沙箱可能拒绝 COM 或子进程标准输入输出捕获，私有主目录写入也可能需要逐次调用提权。`PROCESS_ACCESS_DENIED` 仅报告白名单中的 EPERM/EACCES 子进程错误，不转发子进程诊断，也不判定限制来源。受隔离调用实际发生拒绝后，使用现有 shell 工具的 `sandbox_permissions` 和具体 `justification` 重试完全相同的命令；不要全局关闭隔离。通用 `PROCESS_UNAVAILABLE` 要求先诊断再重试，不能证明访问遭拒。xlflow 的 recovery、dirty 和 error 字段仍然是权威信息。附加和外部会话分离可能清理临时运行时组件；创建按钮和覆盖表单可能保存工作簿。辅助程序不会在初始化期间执行宏。

在 Windows 上从仓库运行 `node --test excel/workbook-binding.test.mjs`。这些测试使用模拟的 COM 和 CLI 操作，验证真实本地文件和格式判断，包括未保存状态、歧义、无法访问的进程、另存为、项目/源码/会话保留，以及附加和拉取顺序。原生只读冒烟检查也枚举了当前打开的 Excel 工作簿。原生多实例附加/拉取，以及原始工作簿的事件和视觉行为，需要指定的临时工作簿证据；测试不声称已完成这些检查或安装程序交付。

## 开发说明

无。
