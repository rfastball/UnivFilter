# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

`UnivFilter` — Excel VBA add-in that adds a custom Universal Filter (in-cell button + context menu + global shortcuts) to any ListObject/AutoFilter range. The shipping artifact is `../UnivFilter(v2.6).xlsm` (migrated from `.xlsb`); the files under `src/` are VBA exports of that workbook's modules.

## Build / Test Workflow

The `_VbaHarness` dev loop is retired (no longer used on any project). Round-trip goes through `CodeManager.bas`: `ExportAllVBAElements` (xlsm → timestamped export folder) and `ImportVBA_From_ImportFolder` (files in `ThisWorkbook.Path & "\IMPORT\"` → xlsm; `.frm` requires its `.frx` next to it). Requires "Trust access to the VBA project object model" in Excel Trust Center.

Tests assert via `LogShim.AssertEq` (no raise on failure). `Test_UnivFilterHarness.Test_UnivFilter_SpecValidation` wraps `modTester.Run_SpecValidation`, which builds spec/report sheets inside the workbook and exercises end-to-end filter scenarios. `LogShim.bas` is a plain project-owned test helper now (not harness-managed) — edit as needed.

## Core_* modules are a deployed package — edit upstream

`Core_Enums`, `Core_modNorm`, `Core_modCompare`, `Core_clsSet` are the shared **UnivCore** package. Their editing home is the sibling `../../UnivCore` project; they are deployed here via `core.manifest` + `CodeManager.Pkg_ExportCorePackage`. Local edits to `Core_*` files will be overwritten on the next package sync — make Core changes in UnivCore and re-deploy.

## Architecture

### Two entry paths share one pipeline

```
[Button click / context menu]        [Ctrl+Shift+Q / Ctrl+Shift+U]
   modUnivFilter.ShowUniversalFilter      modShortcuts.UnivFilter_*Handler
              │                                       │
              └──────────────► frmFilter01 ◄──────────┘
                                    │
                                    ▼
                  modProcessManager.ExecuteFilterProcess
                      │              │              │
                      ▼              ▼              ▼
              MatchingEngine   ResultHandler   (Missing → Clipboard)
```

- `modUnivFilter.ShowUniversalFilter` is invoked by per-column shape buttons (`SmartFilterBtn_<addr>`, `.OnAction` set in `ToggleSearchButtonsInSelection`) or by the cell context-menu entry registered in `SetupContextMenu`. It detects whether the target column lives in a `ListObject` or `AutoFilter`, injects the range into `frmFilter01`, and registers the form in `G_OpenForms` keyed by `sheetName!columnIndex` to prevent duplicates.
- `modShortcuts.UnivFilter_QuickHandler` (Ctrl+Shift+Q) reads clipboard text and runs the filter without showing UI **unless** the column's `AssetType` is `HYBRID`/`ERROR`/`DATE`, in which case it falls back to showing the form with the clipboard pre-filled.
- Shortcuts and context menu are bound in `현재_통합_문서.Workbook_Open` (`Application.OnKey` + `SetupContextMenu`) and torn down in `Workbook_BeforeClose`.

### Column data lifecycle

`modCache.GetColumnAsset(rng)` is the single chokepoint for reading a column. It returns a `clsColumnAsset` (data `Core_clsSet` + format string + detected type). Cache key is `rng.Address(External:=True)` so workbook/sheet/address uniquely identify an asset.

Cache invalidation is **event-driven**, not time-based: `현재_통합_문서.AppEvents_SheetChange` and `AppEvents_SheetCalculate` call `modCache.InvalidateCacheBySheet` on any sheet edit/recalc outside this workbook. If you add a code path that mutates user data, do not manually invalidate — let the app events do it.

The "true" last row uses `max(End(xlUp), UsedRange.LastRow)` (see `modCache.GetColumnAsset` ~line 432). This is a deliberate defense against `End(xlUp)` under-reading filtered ranges; do not "simplify" it back to a single `End(xlUp)`.

### Type detection drives UI lock state

`modFilterUtils.GetColumnIdentity` classifies each column into `FilterDataType` (NUMBER, STRING, DATE, PERCENT, ROUNDED, PERCENT_ROUNDED, HYBRID, ERROR). The detected type rides along on `clsColumnAsset.DetectedType` and `frmFilter01` uses it on `StartInitialization` to decide whether to call `EnforceHybridMode` / `EnforceRoundingMode` / `EnforceDateMode`, which **lock UI controls** and force partial-text-match. The matching engine itself (`modMatchingEngine.MapIdentityToEnum` + `ExecuteValueStrategy` / `ExecuteStringStrategy`) branches on the same type.

When changing filter behavior, update **all three layers** (detection in `modFilterUtils`, UI enforcement in `frmFilter01.Enforce*Mode`, strategy selection in `modMatchingEngine`) or the layers will disagree.

### Global state lives in `modUnivFilter`

`G_LastExcludeMode` / `G_LastStackMode` / `G_LastMatchMode` persist UI toggles across form instances within the session. `G_OpenForms` is the form-instance registry. These are not redundant with form-local state — they are session memory.

### Module dependency map

```
현재_통합_문서 (ThisWorkbook) ─► modCache, modUnivFilter (via shortcut bindings)
modUnivFilter ─► frmFilter01
frmFilter01 ─► modCache, clsColumnAsset, FocusAPI, modProcessManager
modProcessManager ─► modMatchingEngine, modResultHandler, ClipboardAPI, clsColumnAsset
modMatchingEngine ─► modFilterUtils, Core_clsSet, clsColumnAsset
modResultHandler ─► modFilterUtils
modCache ─► Core_clsSet, clsColumnAsset, modFilterUtils
modShortcuts ─► modUnivFilter, frmFilter01, ClipboardAPI
clsColumnAsset ─► Core_clsSet
Core_clsSet ─► Core_Enums          (UnivCore package)
Core_modCompare ─► Core_modNorm    (UnivCore package)
Test_UnivFilterHarness ─► modTester, LogShim
```

`modLogger.Log` is called from nearly every module but is **off by default** — `ENABLE_ALL_LOGGING = False` in `modLogger.bas`. Flip it to True during debugging; remember to flip back before export (the constant ships in the xlsm).

## Conventions specific to this codebase

- All comments, log messages, and user-facing strings are in **Korean**. Keep that language when editing; mixing in English breaks tone consistency with the existing code.
- Module headers and procedure `@desc` follow the schema in the parent `VBA_Projects/CLAUDE.md` §2 — apply it to new Public Sub/Function declarations.
- No em dash (U+2014) anywhere in source, comments included — it corrupts on the import path. Use `ChrW(&H2014)` if the character itself is needed.
- Forms are **modeless** (`.Show vbModeless`) so Excel stays interactive. Don't switch to modal without understanding the `G_OpenForms` re-activation path.
- `Application.Run` and `OnAction`-via-string indirection: shape buttons reference `ShowUniversalFilter`, context menu items reference `CallFilterFromMenu` / `ToggleSearchButtonsInSelection` / etc. by string. Renaming any of those Public Subs requires updating the corresponding `.OnAction =` assignments in `modUnivFilter`.
- VBA `Public` consts/enums declared in any module are project-global — see the "Global Definitions" section of `../UnivFilter(v2.6).md` for the current inventory before adding new ones.

## Reference document

`../UnivFilter(v2.6).md` is an auto-generated full-codebase snapshot (dependency graph, API surface, every module's source; `../UnivFilter.md` is an older snapshot of the same kind). When you need a quick API lookup across modules, grep the snapshot rather than opening every `.bas`. It can drift from `src/` — when they disagree, trust `src/`.
