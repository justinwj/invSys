# Slice 12 Reviewed Code-Bloat Cleanup Evidence

- Date: 2026-07-27
- Contract: delete only reviewed, protected runtime code while preserving the
  consolidated role behavior established by Slices 5-11
- Result: GREEN

## D13 RED

- `Test-Slice12ReviewedCleanup.ps1` initially passed 3 of 11 checks. The eight
  expected failures proved that the reviewed runtime diagnostics and test
  module were still packaged, the two superseded standalone Shipping forms
  remained, the 21 scanner-HIGH unreachable procedures remained, and the
  component/procedure/candidate/duplicate metrics had not improved.
- The protecting Phase 6 Boxing behavior range 124-126 passed 3 of 3 before
  implementation. This locked the archive filter, main Shipping form
  initialization, and projected component-inventory behavior before removing
  the superseded standalone forms.

The focused RED was behavioral and structural: the deletion manifest named
every intended removal, its scanner basis, and its protecting test. No compile
failure, missing fixture, or unavailable workbook was treated as RED.

## GREEN

| Evidence | Result |
|---|---:|
| `Test-Slice12ReviewedCleanup.ps1` | 11/11 |
| Phase 6 protected Boxing range 124-126 | 3/3 |
| Slice 11 Shipping/Boxing regression | 11/11 |
| Slice 5 behavior locks | 13/13 |
| Slice 6 shadow contracts | 10/10 |
| Packaged XLAM validation | 60/60 |
| Live packaged role workflow | 46/46 |
| Packaged RibbonX validation | 156/156 |
| Operations shadow validation | 13/13, zero collisions |
| Tool contract regression | 62/62 |
| Maintenance baseline regression | 19/19 |

The two removed Shipping forms were already superseded by the Box Builder and
Box Maker tabs in `frmShipmentsTally`. Their executable projection behavior is
now protected at the typed `modBoxingService` boundary, and the initialization
smoke test exercises the consolidated main Shipping form.

The three developer-only Core modules were moved out of runtime source:
diagnostics now live under `tools/vba/legacy-diagnostics`, and the legacy test
module lives under `tests/legacy-vba`. The cleanup manifest records the exact
relocations and the 21 removed HIGH-confidence private procedures.

## Static ratchets

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| Components | 169 | 164 | -5 |
| Procedures | 4690 | 4526 | -164 |
| Maintenance candidates | 1033 | 965 | -68 |
| Duplicate-body groups | 199 | 187 | -12 |
| Oversized-module ratchets | 24 | 23 | -1 |
| Literal `Application.Run` targets | 9 | 9 | 0 |
| Unresolved dynamic calls | 48 | 48 | 0 |

The reviewed dynamic-root hash is unchanged, so no exception was added to
silence the scanner. Generated maintenance reports retain the remaining legacy
findings for later review instead of deleting them without reachability proof.

## Additional diagnostic run

The broader Phase 6 range 1-50 passed 47 of 50. The three failures are the same
authentication/session tests already present in the pre-Slice-12 generated
baseline and do not touch the reviewed cleanup surface. The isolated range
20-24 reproduced them without a compile failure. They are not used as Slice 12
GREEN evidence and remain visible for the final Release 1 reconciliation gate.
