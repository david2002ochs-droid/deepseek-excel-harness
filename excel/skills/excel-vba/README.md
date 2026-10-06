---
description: Guarded live workbook binding and private xlflow source projects for the Excel pane VBA skill.
kind: package-library
---

# Excel VBA helper reference

English | [中文](README.zh.md)

## Summary

The [skill](SKILL.md) uses [workbook.mjs](scripts/workbook.mjs) to validate the pane's captured local locator and operate on one open original. [inventory.ps1](scripts/inventory.ps1) reads workbook identity through native Windows handles and COM in every Excel process. Neither inventory nor validation opens, closes, saves or changes workbook/security settings. The product copies this whole skill directory into `<home>/skills/excel-vba`; the helper uses `<home>/runtime/xlflow-0.35.0` and checks the CLI/bridge hashes before execution.

## Research guidance

The skill's [targeted research guidance](SKILL.md#research-when-needed) calls for Microsoft Learn lookups when Excel/VBA details remain uncertain. It uses available tools and relevant existing references; the skill does not configure a Learn connection or replace workbook behavior checks.

## Requests and source ownership

Run `node "<installed skill>/scripts/workbook.mjs" "<request.json>"`. The request contains `action`, the current request's pane `context`, and an `args` string array for commands. Context uses `{workbook:{id,path},worksheet:{id,name},selection:{workbookId,address},observedAt}`. Success returns `ok`, target identity and, when applicable, `project`, `exitCode` and the complete xlflow JSON `output`. A refusal returns `ok:false` with `error.code` and corrective `error.message`; process exit is nonzero. The module exports `validateBinding(context, operations)` and `workbookRequest(request, operations)` for integration/tests; production callers may supply `operations.home` and otherwise keep native operations.

| Action | Behavior |
| --- | --- |
| `validate` | Local file/format checks and complete live inventory; no writes or xlflow invocation. |
| `prepare` | Atomically creates a private project or checks an existing managed project; no attach or pull. |
| `bootstrap` | Creates a new project, revalidates before attach, then revalidates before Excel/session pull. Never overwrites an existing project. |
| `rebind` | Explicitly updates helper-owned pane binding after validation; preserves source/configuration/session and refuses a conflicting session. |
| `command` | Revalidates before invoking supported source/live commands. Source `inspect calls/symbols` omits session checks/flags; workbook observations require the external session. Arguments stay separate process arguments; target overrides, file backend and managed lifecycle are refused. |

Projects live under `<home>/vba-projects/<path-hash>`, with exact `[excel].path`, real `src/{modules,classes,forms,workbook}` directories, sidecar UserForm mode and a helper-owned `binding.json`. Creation uses a sibling staging directory and atomic rename. Reuse requires the managed configuration and the same workbook/process; customized or unrelated configurations fail without replacement. Serialize helper calls for one workbook. Commands that require a session must match `.xlflow/session.json` path, PID and external owner. Existing sessions are resumed, never silently reattached. Pull only accepts empty source directories, so subsequent source reconciliation requires an explicit authority decision outside automatic bootstrap.

The guard rejects `--input` and `--save-as` in split or equals form before inventory. Among supported commands, only `export-image` accepts a workbook positional argument; its documented flag arities distinguish option values from that refused position. Macro arguments and source paths are not rejected by filename suffix. `run --save` remains an explicit request to persist the original after successful execution.

## Limits and verification

Only `.xlsm`, `.xlsb`, `.xlam` and `.xltm` local drive locators are accepted, with matching live `FileFormat`. Unsaved edits in an already-saved original remain valid. Missing, cloud, network or unsaved-name locators, duplicate matches, inaccessible Excel processes and changing process inventories fail. Windows native object model access must be available; invisible/inaccessible instances cannot be assumed empty. A captured locator cannot establish current-pane identity after Save As, even if the old path still matches another open workbook. Request fresh pane metadata when identity is uncertain; no TTL substitutes for it. User actions between inventory and CLI execution remain a race.

The helper sends the local inventory script as a static encoded PowerShell command without changing execution policy. The existing shell sandbox may deny COM or child stdio capture, and private-home writes may require per-call escalation. `PROCESS_ACCESS_DENIED` reports only whitelisted EPERM/EACCES child-process errors, without forwarding child diagnostics or identifying the restriction's source. After an actual denial in a confined call, retry the exact command with the existing shell tool's `sandbox_permissions` and a specific `justification`; do not disable confinement globally. Generic `PROCESS_UNAVAILABLE` requires diagnosis before retry and does not establish denial. xlflow recovery/dirty/error fields remain authoritative. Attach and external detach can clean transient runtime artifacts; button creation and form overwrite may save. The helper does not execute macros during bootstrap.

Run `node --test excel/workbook-binding.test.mjs` from the repository on Windows. These tests exercise real local file/format decisions with fake COM/CLI operations, including unsaved state, ambiguity, inaccessible processes, Save As, project/source/session preservation and attach/pull sequencing. A read-only native smoke also enumerated the current open Excel workbook. Native multiple-instance attach/pull and original-workbook event/visual behavior require designated disposable workbook evidence; the tests do not claim those checks or installer completion.

## Dev Note

None.
