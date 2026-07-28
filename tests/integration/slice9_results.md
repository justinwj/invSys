# Slice 9 Production Layout Standardization

Date: 2026-07-27

## Scope

Slice 9 replaces Production's one-off coordinate arithmetic with a
Windows-aware declarative anchor layout:

- the form declares minimum, default, and expanded geometry in `@FormLayout`;
- all four pages register controls through an Operations-local typed anchor
  manager;
- native minimize, restore, maximize, and resize behavior remains available;
- high-DPI displays apply a native-window DPI zoom before the anchor baseline
  is captured; and
- packaged runtime diagnostics verify the same form surface used by operators.

The accepted geometry is 1110x690 points for minimum/default and 1350x750
points for the expanded validation case. Maximized layout is validated
separately against the current Windows working area.

## D13 trace

`tests/tooling/Test-Slice9ProductionLayout.ps1` was created and observed
meaningfully RED before implementation. Its initial result was 1 passed and
7 failed because Production had no declarative anchor manager or layout
metadata and still resized controls through page-specific coordinate
arithmetic.

Additional focused RED states discovered without weakening the assertions
were:

- form-level `Controls` enumeration included nested controls, so geometry
  checks initially compared unrelated coordinate systems;
- MSForms exposes a page control's coordinate owner through `Container`, not
  `Parent`;
- a `Page` does not expose the usable canvas dimensions expected by the
  generic anchor engine and required normalization through its `MultiPage`;
- list-box runtime sizing produced a one-point overlap on the loader page;
- the original screenshot/window lookup could hang when selecting the native
  form window;
- geometry inspection changed form dimensions while maximized and could hang;
- logical point-space checks passed while visual inspection exposed 150%
  Windows-DPI clipping; and
- copied width/height helpers increased the duplicate-body baseline from 190
  to 192.

The repairs use immediate-container geometry, normalized page extents,
process-owned native-window capture, read-only current-geometry inspection,
DPI-aware zoom and baseline recapture, and one axis-parameterized helper. The
final focused result is 8 passed and 0 failed.

## GREEN and regression evidence

| Evidence | Result |
|---|---|
| `tests/tooling/Test-Slice9ProductionLayout.ps1` | 8 passed, 0 failed |
| `tools/validate_slice9_production_layout.ps1` | PASS: 3 sizes x 4 pages |
| Packaged maximized geometry | PASS: 4 pages |
| Packaged native window transitions | PASS: minimize/restore/maximize/restore |
| Packaged geometry failures | 0 out-of-bounds, 0 interactive overlaps |
| Source harness tests 189-192 | 4 passed, 0 failed |
| Source harness tests 212-252 | 41 passed, 0 failed |
| `tools/validate_phase6_packaged_xlams.ps1` | 59 passed, 0 failed |
| `tools/validate-operations-shadow.ps1` | 13 passed, 0 failed |
| Shadow collision groups | 0 unresolved |
| `tests/tooling/Test-Slice0ToolContracts.ps1 -Mode All` | 62 passed, 0 failed |
| `tests/tooling/Test-Slice3Baseline.ps1` | 19 passed, 0 failed |

The runtime report is
`tests/integration/slice9_layout_results.md`. Its minimum, default, and
expanded screenshots are under `tests/integration/slice9-layout/` and were
visually inspected. Each shows all page content, the action row, status
surface, and Close action without clipping or overlap.

Every Production row in the broader 45-check live workflow passed, including
the two-consecutive-batch captured-workbook action. The 11 remaining
Receiving, Shipping/Boxing, and downstream aggregate failures belong to later
slices. The raw live report contains machine/runtime paths and is intentionally
not release evidence.

## Static maintenance evidence

| Metric | Slice 8 | Slice 9 | Delta |
|---|---:|---:|---:|
| Packages | 8 | 8 | 0 |
| Components | 162 | 165 | +3 |
| Procedures | 4,578 | 4,611 | +33 |
| Source lines | 101,077 | 101,552 | +475 |
| Maintenance candidates | 1,002 | 1,006 | +4 |
| Duplicate-body candidates | 190 | 190 | 0 |
| Unresolved dynamic calls | 48 | 48 | 0 |
| Literal `Application.Run` calls | 10 | 10 | 0 |

All three new Operations layout components are below the 1,000-line
new-module ratchet: `cOperationsAnchorItem` is 127 lines,
`cOperationsAnchorManager` is 80, and `modOperationsLayout` is 11.

`frmProduction` increased from 4,606 to 4,788 lines. This is the explicit
Slice 9 bloat exception: 182 lines replace imperative resize branches with
declarative page registrations and add the packaged, operator-surface geometry
diagnostic required to prove the layout acceptance gate. `mProduction`
increased by 40 lines for primitive packaged-test adapters and
`modProductionFormWindow` by 35 lines for DPI/window support. Duplicate-body,
dynamic-call, and literal-dispatch metrics did not regress.
