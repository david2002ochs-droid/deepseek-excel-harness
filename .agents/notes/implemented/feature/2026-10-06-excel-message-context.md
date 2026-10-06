# Agent Note: Small Excel message context and bound cell citations

Status: implemented

English | [中文](2026-10-06-excel-message-context.zh.md)

## Problem

Embedding the Harness UI in Excel does not tell the agent which workbook or selection a user means. A selection label shown earlier can also differ from the selection when a later message is sent. Cell citations need a workbook binding so a historical answer cannot navigate a different pane accidentally.

## Decision

The [Excel entry](../../../../excel/README.md) opts the existing conversation plugin into a parent-frame Office bridge. Each ordinary message starts a fresh metadata capture at submission, before asynchronous attachment preparation. Its captured workbook and worksheet identifiers, host workbook name and available file locator, qualified selection address and observation time join the ordinary durable user-message text. The model projection and replay therefore retain the same input; no separate session event format or agent-loop change is needed.

Context JSON is limited to 2048 UTF-8 bytes, plus fixed marker and citation-instruction text. The wrapper reads no cell values, formulas, VBA or complete sheet inventory. A failed or oversized capture prevents admission and restores the draft rather than truncating the selected target. The ordinary standalone Web UI remains independent of this opt-in.

Assistant citations use `[[cite:Sheet1!A1:B2]]`, including Excel quoting for worksheet names. Only bounded, valid A1 references in assistant prose receive selection controls. Code examples, invalid tokens and responses without matching captured metadata remain text. The originating response supplies its workbook binding; a later message cannot replace it. The parent verifies the exact child window, origin, request identifier and pane identifier, validates the range again, then activates the worksheet and selects the range without writing cells.

The workbook identifier is taskpane-local and changes on reload or a changed Office document location, including Save As. Navigation revalidates that location before selecting. It is not a native COM handle. Office may omit a locator for an unsaved workbook; that binding relies on Office keeping the pane attached to its document. Any subsequent COM operation must resolve and verify its own target. This feature adds neither a COM execution tool nor a model route.

## Alternatives considered

**Copy the existing VBA integration's full preparation scan.** Its initial cell, formula and VBA scans exceed the owner's requested small orientation packet. The agent can retrieve task-specific content later through its tools.

**Use a global current selection or workbook for responses.** Asynchronous submissions, steering and historical answers can then borrow a later selection or map. Per-message capture and response-owned citation binding retain the intended target.

**Introduce a new session event or custom external-link scheme.** Ordinary user-message text already supplies durable model input. Cell navigation belongs to an explicit Excel callback, while external-link delegates retain their existing URL rules.

## Consequences

The agent receives small, reconstructable orientation and exact selection references without an initial workbook scan. Citation clicks change the active worksheet and selected range; citation text alone is not evidence of reading or modifying cells. Reloaded or unavailable bindings require a new message with fresh context.

Parent tests cover metadata-only reads, bounds, source/origin checks, overlapping requests, disposal and navigation validation. Client tests cover submission timing, failure restoration, durable message replay, safe prose rendering and response binding. Those tests do not prove a paid model run or native Office acceptance; these require separate evidence. The final Windows installer remains a separate delivery requirement.
