# Slice 11 Shipping and Boxing Stabilization Evidence

- Date: 2026-07-27
- Normative workflow: local staging/locks -> submission -> processor application
  -> NAS/read-model refresh -> exact lock release
- Result: GREEN

## D13 RED

- The initial focused contract run passed 0 of 10 checks. Shipping had no typed
  workflow state, no separated posting/Boxing services, no captured-workbook
  tabbed shell, and separate Ribbon launch paths for Box Builder and Box Maker.
- The first exact release runtime test failed because the release path required
  the prohibited `ROW` field. After resolving by `System_Key`, the next run
  exposed a stale projected-inventory overlay keyed by the former row identity.
- The saved Shipping workflow fixture failed because its shipment
  `System_Key` did not identify the seeded inventory entity. This was corrected
  in the fixture without weakening runtime validation.
- The legacy sent-row tombstone test failed with
  `System_Key ... was not found in invSys`; its fixture used different shipment
  and inventory identities. The corrected fixture retained exact matching.
- The first full live role run stalled on an offscreen success message. The
  validator's dialog dismissal was upgraded to locate and click the actual
  Excel dialog button.
- After the first tab composition GREEN, the strengthened existing-action
  contract passed 10 of 11 checks and failed only because New Box,
  component Add/Remove, Delete Version, Archive Box, and Delete Box were not
  reachable from the Box Builder tab.

## GREEN

| Evidence | Result |
|---|---:|
| `Test-Slice11ShippingBoxingStabilization.ps1` | 11/11 |
| Phase 6 tombstone tests 116-117 | 2/2 |
| Phase 6 exact lock-release test 135 | 1/1 |
| Phase 6 Shipments Sent test 166 | 1/1 |
| Phase 6 restart/stale-state tests 187-188 | 2/2 |
| Packaged XLAM validation | 60/60 |
| Live packaged role workflow | 46/46 |
| Packaged RibbonX validation | 156/156 |
| Operations shadow validation | 13/13, zero collisions |
| Slice 5 behavior locks | 13/13 |
| Slice 6 shadow contracts | 10/10 |
| Tool contract regression | 62/62 |
| Maintenance baseline regression | 19/19 |

The packaged Shipping form opened modelessly against its captured operator
workbook and exposed Shipping, Box Builder, and Box Maker tabs. The Box Builder
tab exposes New Box, component Add/Remove, Save Box, Update Version, New
Version, Delete Version, Archive Box, and Delete Box. The Box Maker tab exposes
Make Boxes and Unbox. The Operations Ribbon contract retains only the main
Shipping launcher.

The live workflow exercised the real Shipments Sent form action, processor
application, read-model publication, exact lock clear, and retained captured
workbook. Focused runtime tests prove exact `System_Key` release, idempotent
replay, tombstone filtering, and no completed-staging resurrection after
restart.

## Static ratchets and scoped exception

| Metric | Slice 10 | Slice 11 | Delta |
|---|---:|---:|---:|
| Components | 166 | 169 | +3 |
| Procedures | 4601 | 4690 | +89 |
| Literal `Application.Run` targets | 9 | 9 | 0 |
| Unresolved dynamic calls | 48 | 48 | 0 |
| Duplicate-body groups | 190 | 199 | +9 |
| `frmShipmentsTally.frm` lines | 1855 | 2525 | +670 |
| `modTS_Shipments.bas` lines | 21657 | 22389 | +732 |

The line-growth and duplicate-body increases are an explicit Slice 11
exception for the specified Shipping tab composition and the temporary
separation of typed state, posting, and Boxing service boundaries before the
reviewed Slice 12 cleanup. All three new runtime components remain below the
1000-line new-module threshold. The nine duplicate groups are reviewed
equivalences: typed workflow accessors, paired tab handlers, retired-launcher
shims, and one managed-table predicate. No late-binding metric increased, and
no duplicate was deleted without the Slice 12 reachability review and
protecting tests.
