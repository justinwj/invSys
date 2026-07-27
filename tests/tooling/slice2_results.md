# Slice 2 runtime-state extractor results

**Slice:** 2 - Runtime-state extractor MVP

**Recorded:** 2026-07-27

**Runtime workbook required:** No

**Operational workbook opened or changed:** No

## D13 trace

The focused Tool B RED was recorded before implementation:

```text
ToolB.EntryPoint: tools/export-invsys-runtime-state.ps1 is absent; this is the
expected Slice 0 RED.
```

The offline comparison requirement received a separate focused RED:

```text
ToolB.Comparison.EntryPoint: tools/compare-invsys-reports.ps1 is absent;
reports cannot be compared offline.
RESULT passed=11 failed=1
```

Focused GREEN:

```powershell
.\tests\tooling\Test-Slice0ToolContracts.ps1 -Mode Runtime
```

```text
RESULT passed=15 failed=0
```

The test runs the extractor twice with a fixed timestamp and proves
byte-identical JSON and Markdown. It also proves redaction, default exclusion
of row values, accurate legacy-package and retired-`ROW` warnings, agreement
between canonical JSON and rendered Markdown, zero mutation counters, and
identical reported before/after hashes.

An independent test-harness SHA-256 check around both invocations proves the
synthetic inspected-workbook fixture remained unchanged. The extractor source
is also rejected if it contains workbook open, save, close, refresh, processor,
repair, Excel-startup, or Excel-quit paths.

## Read-only runtime contract

The runtime schema and redaction policy are frozen at version `1.1.0`. The
report records:

- `FIXTURE`, `LIVE_ATTACHED`, or `NO_SESSION` inspection mode;
- whether Excel was started by the tool, constrained to `false`;
- explicit zero-constrained counters for open, close, save, refresh,
  processor, repair, and all mutating actions;
- before/after SHA-256 and `unchanged` status for inspected files;
- schema/count/status evidence for loaded add-ins, workbooks, managed tables,
  config, queues, processor state, snapshots, staging, forms, and warnings; and
- redaction counts without secret values or unrestricted row-level data.

Live inspection uses `GetActiveObject("Excel.Application")` only. The tool
does not create Excel, open a workbook, call a project macro, save, close,
refresh, repair, process, or quit Excel.

## No-session proof

With no `EXCEL` process present, the extractor ran without a fixture:

```powershell
.\tools\export-invsys-runtime-state.ps1 `
  -OutputDirectory <temporary-directory> `
  -ReportTimestampUtc 2026-07-27T00:00:00Z
```

Observed validated report:

```json
{
  "inspectionMode": "NO_SESSION",
  "excelStartedByTool": false,
  "mutations": 0,
  "loadedAddins": 0,
  "openWorkbooks": 0
}
```

This path validates the generated JSON against the committed schema and
produces Markdown from the same in-memory report.

## Offline comparison

`tools/compare-invsys-reports.ps1` compares two JSON reports without Excel.
The focused test proves that it:

- reports zero differences for semantically identical reports; and
- reports one `VALUE_CHANGED` difference at `$.capturedAtUtc` for a controlled
  changed report.

## Regression

```powershell
.\tests\tooling\Test-Slice0ToolContracts.ps1 -Mode All
```

```text
Contract and fixture checks: 37 passed
Static scanner checks: 10 passed
Runtime extractor checks: 15 passed
RESULT passed=62 failed=0
```

All changed PowerShell files parse successfully. `git diff --check` reports no
whitespace errors.

No VBA runtime module, form, RibbonX package, build script, or operational
workbook changed. Packaged Excel regression and runtime code-size comparison
do not apply to this developer-diagnostics slice.
