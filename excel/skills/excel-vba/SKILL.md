---
name: excel-vba
description: Edit VBA, worksheet macro buttons and UserForms in the original workbook identified by Excel pane context, using the installed xlflow runtime and guarded live binding. Use for source-backed Excel automation; ordinary cell edits do not require VBA initialization.
---

# Excel VBA in the pane workbook

For ordinary values, formulas, ranges and formatting, use the available shell tool for a COM cell operation with separate exact-target validation. The pane supplies metadata rather than worksheet execution tools; do not initialize a VBA project for cell-only work. For VBA, worksheet macro buttons or UserForms, read the installed sibling `../xlflow/SKILL.md` and only the relevant upstream reference before choosing commands. The product installs the pinned runtime and both skills; do not install software, search other profiles or resolve `xlflow` from PATH.

## Bind before attaching

Use [scripts/workbook.mjs](scripts/workbook.mjs) with a JSON request file containing `action`, the current request's `context` and, for `command`, an `args` array. Run `node "<this skill directory>/scripts/workbook.mjs" "<request.json>"` through the existing shell tool. The installed helper derives the private home from its own location and returns structured JSON. [The helper reference](README.md) defines actions and refusal recovery.

1. Start with `action: "validate"`. It reads all Excel processes without attaching or writing. Require the exact saved local drive path, supported macro-enabled format and one unique open match. Saved workbooks with unsaved live edits are supported. Missing, cloud, unsaved-name, network, ambiguous and inaccessible locators require correction; never choose a workbook by name or ActiveWorkbook, convert `.xlsx` implicitly or perform silent Save As.
2. `action: "bootstrap"` creates a private source project, attaches the original and pulls through the Excel/session backend. It never uses `xlflow init`, which copies the workbook. Bootstrap only accepts a new project; `prepare` creates or checks a project without attaching. Existing source is preserved. Establish authority before editing: the initial live pull captures unsaved VBA; later source edits remain authoritative until deliberately imported. A pull into existing source is refused rather than replacing it.
3. Use `action: "command"` and structured arguments for subsequent operations. Every invocation rechecks the live locator, private configuration and external session identity. Use installed CLI help for exact syntax. The helper pins the executable and bridge, supplies JSON/session flags and requires `--no-save` for push/test. It refuses target/backend overrides and managed lifecycle operations. Resume an existing project; use explicit `rebind` with current pane metadata after a pane reload or binding change without replacing its source.

Pane metadata is captured on user submission. Its id is a pane binding, not a COM handle; `observedAt` is not proof of identity. The helper validates the captured locator, not the current pane. If Save As, relocation or current-pane uncertainty intervenes, request a new user submission with fresh pane context before continuing. An old file still open elsewhere can match an old capture. No TTL resolves that ambiguity. Live revalidation also cannot prevent a user switching state during an external command; stop on conflicting xlflow diagnostics.

After an actual denial in a confined shell call, retry the exact helper command through the existing shell tool's per-call `sandbox_permissions` and `justification`. `PROCESS_ACCESS_DENIED` preserves a safe EPERM/EACCES diagnosis without identifying the restriction's source; explain that specific access refusal. `PROCESS_UNAVAILABLE` is a generic process failure: diagnose it before retrying, without inferring denial. Do not disable the product sandbox or change Excel trust/security settings. A failed inventory proves no uniqueness.

## Research when needed

When correctness depends on uncertain Excel/VBA API details, version support or runtime behavior, reuse relevant findings already available in the task. If they are insufficient, use the available Microsoft Learn MCP, selecting operations from the tools and schemas visible in this session. Search for the exact object/member and relevant symptom or version, then fetch the relevant official page; fetch directly when its URL is known. Apply documentation for Excel/VBA and the target environment, and stop when the implementation decision is supported. If Learn is unavailable, consult official documentation through an already available tool or state the unresolved gap.

Consult only relevant existing xlflow references, project knowledge cards or repository examples, including `../xlflow/references/object-model-traps.md` for the listed runtime traps. Translate the finding into concrete workbook preconditions and a focused behavior or event check using the workflow below. Treat undocumented observations as specific to the observed environment, and cite the source when it materially determines the implementation.

## Route the change and its evidence

| Work | Guidance and required evidence |
| --- | --- |
| Worksheet macro button | Use a public, parameterless procedure in a standard module. Verify the actual button's `OnAction` points to that original workbook and procedure. Exercise the button entrypoint and observe its effect; a macro run alone does not prove wiring. Use the pinned `ui button --help` for syntax. |
| UserForm | Load `../xlflow/references/forms.md`. Keep Designer specs and sidecar code authoritative. Match each control's actual name and event signature, such as `CommandButton1_Click`; exercise the intended event and verify its result. |
| Dialog or headless UI | Load `../xlflow/references/xlflow-ui.md` for dialog policy. Keep GUI entrypoints thin and business logic testable without displaying forms. |
| Business logic or event behavior | Load `../xlflow/references/testing.md`; run focused logic checks and the actual event path. Lint/analyze and successful import do not prove event execution. |
| Rendered appearance | Export and inspect the relevant worksheet/form image. Form runtime/image inspection may use a temporary copy. Report visual evidence separately from original-workbook event proof. |
| Recovery or dirty-state conflict | Load `../xlflow/references/recovery.md` and follow structured recovery instructions. Do not save or blindly retry while recovery is required. |

## Finish deliberately

Validate source, import with `push --fast --no-save`, run focused behavior checks and inspect live outputs. Explain any command's persistence effect before using it: attach removes stale runtime artifacts; worksheet button creation and form overwrite can save even when earlier pushes did not. Do not promise universal no-save behavior. Save only when the verified changes should persist. A validated external `session stop` detaches without closing or saving the original, but may remove transient runtime artifacts; detaching is optional. Do not require saving or closing the user's workbook merely to finish.

Report the original path, source authority, changed artifacts, validation, business/event results, visual evidence when relevant, and whether changes persisted or remain unsaved. Never claim a functional button or form from source review or an image alone.
