# Slice 3 baseline and reviewed-backlog results

**Slice:** 3 - Baseline, root registry, and reviewed cleanup backlog

**Recorded:** 2026-07-27

**Runtime workbook required:** No

**Operational workbook opened or changed:** No

## D13 trace

Focused RED:

```powershell
.\tests\tooling\Test-Slice3Baseline.ps1
```

```text
FAIL Slice3.EntryPoint - tools/create-maintenance-baseline.ps1 is absent;
this is the expected Slice 3 RED.
RESULT passed=0 failed=1
```

Focused GREEN:

```text
RESULT passed=19 failed=0
```

The GREEN run regenerated all six committed artifacts from source with a fixed
timestamp and proved byte-for-byte equality. It also validated the reviewed
backlog schema, separate role/shared-package workstreams, zero automatic or
approved deletions, protecting-test discipline for HIGH-confidence deletion
candidates, oversized-module ratchets, no-growth dynamic-call/duplicate
defaults, and the class-event root convention.

## Complete baseline

The baseline records observed pre-D12 implementation reality:

| Metric | Value |
|---|---:|
| Current build packages | 7 |
| Components | 151 |
| Procedures | 4,441 |
| Source lines | 98,836 |
| Scanner candidates | 1,038 |
| Reviewed/manual backlog items | 1,040 |
| Dynamic roots | 768 |
| Class/WithEvents roots | 31 |
| Oversized-module ratchets | 25 |

The reviewed workstreams are separately recorded:

| Workstream | Items |
|---|---:|
| Receiving | 39 |
| Production | 165 |
| Shipping/Boxing | 174 |
| Shared Operations | 68 |
| Core | 313 |
| Domains | 63 |
| Admin | 218 |

Two manual architecture items retain traceability beyond scanner heuristics:

- D14 greenfield warehouse generation and demo seeding must not call old
  business-inventory import or build `ROW`-to-`System_Key` mapping.
- D12 role sources must move to `invSys.Operations.xlam` without becoming one
  monolithic module.

## Root-registry review

The first scanner baseline exposed class lifecycle and `WithEvents` callbacks
as false HIGH-confidence removal candidates. The root contract now includes
`CLASS_EVENT` for `.cls` procedures such as `Class_Initialize`,
`Class_Terminate`, and control-event handlers.

The regenerated baseline retains 31 such procedures as dynamic roots. No
source procedure was deleted or rewritten.

## Deletion policy and candidate review

The committed backlog preserves original scanner type and confidence
separately from reviewed confidence and disposition.

- Scanner confidence is explicitly not deletion authority.
- Automatic deletion and approved deletion counts are both zero.
- A scanner-HIGH item without a protecting test is downgraded to reviewed
  MEDIUM and marked `REQUIRES_PROTECTING_TEST`.
- Any future reviewed-HIGH `REMOVE`, `REPLACE_DUPLICATE`, or
  `ISOLATE_LEGACY_IMPORT` item must have a reason and protecting test.
- `modExportImportAll` candidates are classified `MOVE_TO_TESTS`, because the
  module is developer export/import support embedded in runtime Core.

The reviewed disposition counts are:

| Disposition | Count |
|---|---:|
| Retain dynamic root | 393 |
| Unresolved/manual investigation | 288 |
| Replace duplicate | 193 |
| Remove, test required | 108 |
| Split module | 36 |
| Replace same-project late binding | 14 |
| Move to tests | 7 |
| Isolate legacy import boundary | 1 |

## Growth ratchets

New modules are limited to 1,000 lines and new procedures to 200 lines.
Existing modules already over 1,000 lines may not grow without an explicit
exception. The largest recorded baselines are:

| Module | Lines |
|---|---:|
| `src/Shipping/Modules/modTS_Shipments.bas` | 21,625 |
| `src/Production/Modules/mProduction.bas` | 12,927 |
| `src/Production/Forms/frmProduction.frm` | 4,452 |
| `src/Core/Modules/modRoleEventWriter.bas` | 2,753 |
| `src/InventoryDomain/Modules/modInventoryApply.bas` | 2,183 |
| `src/Receiving/Modules/modTS_Received.bas` | 2,144 |

Same-project `Application.Run`, unresolved dynamic-call, and duplicate-body
counts may not increase.

## Artifact hashes

| Artifact | SHA-256 |
|---|---|
| `implementation-manifest.json` | `68734aede560d266c5add0720bf18b63fab34d24b365c70693d16022dce0eed9` |
| `implementation-manifest.md` | `4bf01500e0f695dfb2a2838fbfa034f418501448af28a2e7270621f357f63344` |
| `maintenance-candidates.json` | `7dd299d1e31edade13c8ec2cbd1f59c3577de473e019191f65a36f6e29dd5234` |
| `maintenance-candidates.md` | `71683a4241e79952b6d06be7e6ab1ea04fd01cc0c440aa847d81f86aacf1c052` |
| `reviewed-cleanup-backlog.json` | `08a396927db0ad68d0e7ce25966d4dab971d11a337cb8bbc82f1959d6dc1a788` |
| `reviewed-cleanup-backlog.md` | `f6b68ee3cd5c61099aa50685530582c943b3713f38966e336139e8b603b09ea0` |

## Regression

The Slice 1 static-tool contract remains GREEN:

```text
RESULT passed=10 failed=0
```

No VBA runtime module, form, RibbonX package, build map, or operational
workbook changed. Excel compile and packaged action tests do not apply to this
developer-tooling and generated-evidence slice.
