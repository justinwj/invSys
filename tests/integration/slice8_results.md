# Slice 8 Production Controller and Legacy Retirement

Date: 2026-07-27

## Scope

Slice 8 makes the modeless Production form resolve every action through its
captured operator-workbook context and retires Production-specific legacy
authority paths:

- managed identity is `System_Key`; Production contains no managed `ROW`
  header, alias, display field, payload field, or numeric-row authority helper;
- Designs-enabled Production returns released Designs Domain recipes without
  falling through to legacy recipe storage;
- form-to-controller calls inside Operations are direct typed calls;
- Production crosses the Core boundary only through declared primitive,
  array, or JSON contracts;
- Production payload dictionaries and JSON are created inside Operations;
- canonical inventory is changed only through queued events and processor
  application;
- controller test fixtures are explicitly marked and stripped from deployed
  binaries; and
- the captured form context fails closed when its workbook closes rather than
  rebinding to `ActiveWorkbook`.

## D13 trace

`tests/tooling/Test-Slice8ProductionRetirement.ps1` was created and observed
meaningfully RED before implementation. Its initial result was 0 passed and
8 failed because the Production surface still contained `ROW` authority,
same-project dynamic dispatch, direct Core object helpers, legacy mutation
paths, and unisolated runtime test fixtures.

Additional focused RED states discovered during the slice were:

- the canonical inventory picker still resolved the retired `ROW` header;
- a closed captured Production workbook silently rebound the form to a newly
  active workbook;
- Production passed Collections, forms, `WarehouseTarget`, `Workbook`,
  `Worksheet`, and payload dictionary objects across the Core XLAM boundary;
- deployed `mProduction` still contained 18 public test procedures because
  standard BAS imports bypassed the test-only stripper;
- the maintenance baseline rose from 193 to 194 duplicate-body candidates
  after copied JSON/window helpers were introduced;
- the Operations shadow collision report found two public-procedure
  collisions because conditional Mac/Windows branches declared duplicate
  public window procedures; and
- the sequential Production picker range exposed test-context contamination
  that did not reproduce when the test ran alone.

Each RED was repaired without weakening its assertion. The final focused
static result is 14 passed and 0 failed.

## GREEN and regression evidence

| Evidence | Result |
|---|---|
| `tests/tooling/Test-Slice8ProductionRetirement.ps1` | 14 passed, 0 failed |
| Source harness tests 189-192 | 4 passed, 0 failed |
| Source harness tests 212-252 | 41 passed, 0 failed |
| Production session/service tests 244-252 | 9 passed, 0 failed |
| Packaged Production CheckIn/complete/process/log workflow | PASS |
| Packaged two-consecutive-batch captured-workbook action | PASS |
| `tools/validate_phase6_packaged_xlams.ps1` | 59 passed, 0 failed |
| `tools/validate-operations-shadow.ps1` | 13 passed, 0 failed |
| Shadow collision groups | 0 unresolved |
| `tests/tooling/Test-Slice0ToolContracts.ps1 -Mode All` | 62 passed, 0 failed |
| `tests/tooling/Test-Slice3Baseline.ps1` | 19 passed, 0 failed |

The final live Production row returned
`OK|Batches=2|BoundWorkbook=<Production operator workbook>`. Every Production
row in that broader 45-check run passed. Its 11 remaining Receiving,
Shipping/Boxing, and downstream aggregate failures are outside Slice 8 and are
the protected work of later slices. The raw live report contains machine and
runtime paths and is intentionally not release evidence.

## Deployed-runtime inspection

The rebuilt `deploy/current/invSys.Production.xlam` was inspected through its
compiled VBA project:

| Check | Result |
|---|---:|
| `mProduction` compiled lines | 12,006 |
| Public test procedures in deployed `mProduction` | 0 |
| `Application.Run` sites in deployed `mProduction` | 0 |
| `frmProduction` compiled lines | 4,592 |
| `Application.Run` sites in deployed `frmProduction` | 0 |
| `Application.ActiveWorkbook` sites in deployed `frmProduction` | 0 |

Packaged form-action adapters remain available because the D13 two-batch test
must exercise the real handlers. Embedded controller fixture builders are
stripped.

## Static maintenance evidence

| Metric | Slice 7 | Slice 8 | Delta |
|---|---:|---:|---:|
| Packages | 8 | 8 | 0 |
| Components | 159 | 162 | +3 |
| Procedures | 4,569 | 4,578 | +9 |
| Source lines | 101,170 | 101,077 | -93 |
| Maintenance candidates | 1,048 | 1,002 | -46 |
| Duplicate-body candidates | 193 | 190 | -3 |
| Unresolved dynamic calls | 59 | 48 | -11 |
| Literal `Application.Run` calls | 14 | 10 | -4 |

The three new components are the 155-line Core primitive bridge and the
Production-local JSON and window adapters. All are below the 1,000-line
new-module ratchet.

| Runtime module | Previous baseline | Slice 8 | Delta |
|---|---:|---:|---:|
| `frmProduction` | 4,638 | 4,606 | -32 |
| `mProduction` | 13,082 | 12,810 | -272 |
| `modProductionEventCreator` | 317 | 133 | -184 |

No targeted runtime module exceeds its previous bloat baseline.
