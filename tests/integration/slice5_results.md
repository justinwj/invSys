# Slice 5 Pre-refactor Packaged Behavior Locks

Date: 2026-07-27

## Scope

These tests lock the operator action boundaries required before service
extraction:

- Receiving Confirm Writes;
- Production selection, Apply, Check In, Complete Run, and Next Batch across
  two consecutive batches;
- Shipping Shipments Sent;
- captured-workbook behavior after another workbook is activated;
- modeless role launchers;
- the Receiving Purchasing stub; and
- the unified Operations Shipping launcher and tabbed Shipping shell.

The form test seams invoke the existing private click handlers. They do not
substitute direct service calls for operator actions.

## Meaningful RED

`tools/validate_phase6_live_role_workflows.ps1` ran against the packaged XLAMs
with a second operator workbook activated before each tested action.

| Boundary | Observed RED | Layer isolation |
|---|---|---|
| Receiving Confirm Writes | The form remained bound to the Receiving workbook, but the click handler returned `Succeeded=False` with an object-state error; staging remained and no inbox row was submitted. | Runtime target, authentication, and `RECEIVE_POST` capability checks passed. Processor later reported zero applicable rows, proving the failure occurred before Domain processing. |
| Shipping Shipments Sent | The form remained bound to the Shipping workbook, but the click handler reported that the keyed inventory read model had no rows; shipment staging remained and no inbox row was submitted. | Runtime target, authentication, and `SHIP_POST` capability checks passed. Independent Shipping Hold send/return behavior remained green. |
| Production two batches | The exact selection and Apply handlers ran, then the first Check In failed because the controller still considered the test ingredient 0% allocated; Complete Run, Next Batch, and batch two were therefore not reached. | Runtime target, authentication, and `PROD_POST` capability checks passed. No Production inbox event was submitted, isolating the failure above the processor/Domain boundary. |

The packaged live run completed with 26 passing checks and 19 failures. The
additional failures are downstream consequences of the three pre-submit
controller failures or existing `ROW`-addressed Shipping/Production paths; the
raw report is deliberately not committed because it contains machine, user,
and temporary runtime details.

`tests/tooling/Test-Slice5BehaviorLocks.ps1` recorded 7/13 passing source
locks and the following six expected RED gaps:

- Production launcher is not modeless;
- Shipping launcher is not modeless;
- Receiving has no selectable, visibly non-operational Purchasing tab;
- Shipping has no single tabbed shell containing Box Builder and Box Maker;
- `invSys.Operations.xlam` is not yet defined; and
- the single Operations Shipping Ribbon launcher cannot yet be proven.

## Characterization and regression evidence

- Required form-handler wiring: 3/3 PASS.
- Explicit captured-workbook state in Receiving, Production, and Shipping
  forms: 3/3 PASS.
- Receiving modeless launcher: PASS.
- Saved Shipping refresh and close/reopen/process/refresh preservation:
  2/2 PASS (`run_phase6_excel_validation.ps1`, tests 113-114).
- Packaged XLAM validation after adding the action test seams: 59/59 PASS.
- All seven XLAMs rebuilt successfully from the recorded source.
- Slice 0 tooling contracts: 62/62 PASS.
- Slice 3 deterministic maintenance baseline: 19/19 PASS.

The RED tests now identify UI/controller gaps separately from processor and
Domain failures and protect the currently successful packaged behavior during
Slices 6-9.

## Static maintenance evidence

The committed maintenance baseline was regenerated after the final action-seam
implementation and reproduced byte-for-byte.

| Metric | Slice 4 | Slice 5 | Delta |
|---|---:|---:|---:|
| Components | 152 | 152 | 0 |
| Procedures | 4,455 | 4,462 | +7 |
| Dynamic roots | 768 | 768 | 0 |
| Scanner warnings | 42 | 42 | 0 |
| Maintenance candidates | 1,041 | 1,043 | +2 |
| Duplicate-body candidates | 193 | 193 | 0 |
| Unresolved dynamic calls | 59 | 59 | 0 |
| Literal `Application.Run` calls | 14 | 14 | 0 |

Explicit Slice 5 growth exception: the six existing form/controller modules
grew only to expose test entry points that invoke the exact private operator
handlers and report captured-workbook identity. The largest addition is 85
lines in `frmProduction` for the required two-consecutive-batch action
sequence. No production algorithm was duplicated, no new package or runtime
component was introduced, and the duplicate-body and dynamic-call ratchets did
not grow. Slices 6-9 remain responsible for extracting the protected behavior
from these oversized legacy modules.
