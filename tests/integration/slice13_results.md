# Slice 13 D12 Final Operations Cutover Evidence

- Date: 2026-07-27
- Contract: deploy Receiving, Production, and Shipping as one independently
  gated `invSys.Operations.xlam` within the exact five-package Release 1 set
- Result: GREEN

## D13 RED

- `Test-Slice13OperationsCutover.ps1` initially passed 0 of 12 checks. The
  existing build still defined and deployed three standalone role projects,
  had no deployed Operations package or unified Ribbon, and active tooling and
  Admin package lists still required the legacy binaries.
- After the source/tooling cutover, the focused test passed 11 of 13. The two
  expected remaining failures proved that `deploy/current` had not yet been
  cut over: it still contained seven XLAMs and the three standalone roles.
- The coexistence assertion was added and observed RED before the startup
  diagnostic was implemented.
- The version-coherent manifest assertion was added and observed RED before
  build publication generated the hash-verified `R1-5` manifest.
- The historical Slice 6 and Slice 9 contracts then failed 2 of 10 and 1 of 8
  respectively because they still required the temporary pre-cutover package
  state. Their durable shadow and Production-layout assertions were preserved
  while the deployment assertions were advanced to D12.
- The Slice 5 and Slice 11 Ribbon locks each failed only their legacy Shipping
  button identifier after the one-tab cutover. They were advanced to the
  consolidated Operations launcher and then returned GREEN.

## GREEN

| Evidence | Result |
|---|---:|
| `Test-Slice13OperationsCutover.ps1` | 14/14 |
| Packaged five-XLAM compile/surface/restart validation | 54/54 |
| Packaged one-tab Operations/Admin RibbonX validation | 136/136 |
| Live consolidated-package role workflow | 46/46 |
| Phase 6 Add-ins publish/manifest range 27-30 | 4/4 |
| Phase 6 protected Boxing range 124-126 | 3/3 |
| Operations shadow validation | 13/13, zero collisions |
| Slice 5 behavior locks | 13/13 |
| Slice 6 shadow contracts | 10/10 |
| Slice 9 Production layout contracts | 8/8 |
| Slice 11 Shipping/Boxing contracts | 11/11 |
| Tool contract regression | 62/62 |
| Maintenance baseline regression | 19/19 |

`deploy/current` contains exactly:

1. `invSys.Core.xlam`
2. `invSys.Inventory.Domain.xlam`
3. `invSys.Designs.Domain.xlam`
4. `invSys.Operations.xlam`
5. `invSys.Admin.xlam`

The three standalone role binaries are absent from `deploy/current` and were
recoverably archived by the build. `addins-manifest.json` records package-set
version `R1-5`, the exact filenames, sizes, and verified SHA-256 hashes.

The Operations Ribbon contains one shared Session group and separate
capability-gated Receiving, Production, and Shipping groups. Receiving exposes
its read-only Purchasing tab through the main form; Shipping exposes Box
Builder and Box Maker through its main form. None has a separate Ribbon
launcher. Generated same-project callbacks use direct typed calls.

Operations startup detects a simultaneously loaded standalone role XLAM,
declines duplicate role registration, and tells the operator to close Excel and
run `tools/register_current_addins.ps1`. Otherwise, each role initializes
through an isolated typed wrapper and reports the failing role without hiding
the loaded Operations package.

## Static ratchets

| Metric | Slice 12 | Slice 13 | Delta |
|---|---:|---:|---:|
| Build definitions | 6 | 6 | 0 |
| Runtime components | 164 | 164 | 0 |
| Procedures | 4526 | 4531 | +5 |
| Maintenance candidates | 965 | 965 | 0 |
| Duplicate-body groups | 187 | 187 | 0 |
| Oversized-module ratchets | 23 | 23 | 0 |
| Literal `Application.Run` targets | 9 | 9 | 0 |
| Unresolved dynamic calls | 48 | 48 | 0 |

The five added procedures are the required Operations startup,
coexistence-report, and three role-isolation wrappers. No bloat, duplicate,
oversized-module, or late-binding metric regressed. The six static build
definitions are the five deployable packages plus the disposable collision
shadow; only five XLAMs are published.
