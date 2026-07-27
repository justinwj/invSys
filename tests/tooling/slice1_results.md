# Slice 1 static-scanner results

**Slice:** 1 — Static scanner MVP

**Recorded:** 2026-07-27 12:39:23 -07:00

**Runtime workbook required:** No

**Operational workbook opened or changed:** No

## D13 trace

Focused RED was recorded in `slice0_results.md` before Tool A existed:

```text
ToolA.EntryPoint: tools/inventory-vba-surface.ps1 is absent; this is the expected Slice 0 RED.
```

Focused GREEN:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\tests\tooling\Test-Slice0ToolContracts.ps1 `
  -Mode Static
```

```text
RESULT passed=10 failed=0
```

The focused test ran Tool A twice against the synthetic source/Ribbon/build
fixtures and proved byte-identical JSON and Markdown. It also proved direct
call discovery, Ribbon and event roots, unresolved dynamic-call reporting,
duplicate normalized bodies, review-only unreachable candidates, and retired
`ROW` detection.

## Real-repository scan

Command:

```powershell
.\tools\inventory-vba-surface.ps1 `
  -SourceRoot .\src `
  -BuildMapPath .\tools\build-xlam.ps1 `
  -RibbonRoot .\tools\build-xlam.ps1 `
  -TestRoot .\tests `
  -RootRegistryPath .\tools\contracts\vba-dynamic-roots.json `
  -OutputDirectory .\tests\tmp\slice1-real-scan `
  -ReportTimestampUtc 2026-07-27T20:00:00Z
```

Observed implementation:

| Metric | Value |
|---|---:|
| Current build packages | 7 |
| Components | 151 |
| Procedures | 4,441 |
| Exported source lines | 98,836 |
| True literal `Application.Run` targets | 14 |
| Unresolved/concatenated `Application.Run` expressions | 59 |
| Duplicate normalized-body groups | 193 |
| Review candidates | 1,038 |
| Retired `ROW` source warnings | 28 |

The seven parsed packages and outputs match the current build map exactly:
Core, Inventory Domain, Designs Domain, Receiving, Production, Shipping, and
Admin. This is observed pre-D12 reality, not the normative final package set.

Root audit:

| Root evidence | Count |
|---|---:|
| Event-like procedures checked | 315 |
| Event-like procedures missing roots | 0 |
| Ribbon callback references checked | 50 |
| Ribbon callback names missing roots | 0 |
| Ribbon root records | 32 |
| Processor-handler roots | 9 |
| Literal cross-XLAM root records | 14 |
| Test-entry roots | 365 |
| Windows callback roots | 2 |

Every true literal macro target has a corresponding cross-XLAM root. String
concatenations such as `"mProduction." & actionName` are classified as
unresolved dynamic expressions and are not guessed.

## Determinism and schema validation

A second full repository scan with the same timestamp produced identical
SHA-256 values:

| Artifact | SHA-256 |
|---|---|
| `implementation-manifest.json` | `9f391cc1bdf16143a9faf48ffcf3a167aec518eb80061165f68c4c4b2a18309f` |
| `implementation-manifest.md` | `6e6cfcb2134da82cccd764b9faf9c1e75f2d51e982dc2a0dd910aee8c0ee053b` |
| `maintenance-candidates.json` | `b0cb96818c7231ccd877ca71391b1dcd30b0ab6328e1d17a010a93dfd88a5bb1` |
| `maintenance-candidates.md` | `172fd6336156ecd2919ceff9b9e138499d9ef44a450f9ff898e9703dd61be587` |

The scanner validates both JSON reports against the committed version
`1.0.0` schemas before rendering Markdown. Failure to satisfy required types,
properties, enums, constants, or local schema references stops the command.

The real reports remain under the ignored `tests/tmp` diagnostic path. Slice 3
will create the reviewed committed baseline and cleanup backlog after Tool B is
available; this Slice 1 evidence records the reproducible hashes without
prematurely treating scanner candidates as deletion authority.

## Regression

```text
Contract fixtures: passed=36 failed=0
Combined tooling suite: passed=46 failed=1
Remaining failure: ToolB.EntryPoint (expected Slice 2 RED)
```

No VBA runtime module, form, RibbonX package, build script, or workbook changed,
so packaged Excel regression and runtime code-size comparison do not apply to
this slice. The scanner records the baseline but never deletes or rewrites
source.
