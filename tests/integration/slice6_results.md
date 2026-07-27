# Slice 6 Shadow Operations Package and Collision Harness

Date: 2026-07-27

## Scope

Slice 6 establishes the future D12 package boundary without deploying or
registering it:

- `invSys.Operations.xlam` is a disposable shadow containing Receiving,
  Production, and Shipping/Boxing source sets;
- Core and both Domain dependencies are rebuilt beside the shadow;
- legacy role XLAMs remain the active `deploy/current` packages;
- the shadow has no RibbonX part and cannot target `deploy/current`;
- one registration-only Operations startup replaces colliding role startup
  entry points inside the shadow; and
- component, public standard-module procedure, and Ribbon callback collisions
  are reported deterministically.

## D13 trace

Initial meaningful RED from
`tests/tooling/Test-Slice6ShadowOperations.ps1` was 2/8:

- no Operations shadow project existed;
- complete-project selection was absent;
- no shadow build entry point existed;
- no collision harness or reviewed resolution contract existed;
- no packaged shadow validator existed; and
- only the required non-deployment and legacy-package-preservation checks
  passed.

Two focused tooling defects were then isolated:

- 8/9 RED: static inventory could not parse multiple source directories in one
  build project;
- 9/10 RED: the first parser repair admitted the combined project but dropped
  the final Admin project block from the generated manifest.

Final focused GREEN is 10/10. The manifest now contains both Admin and
OperationsShadow and inventories all four Operations source roots.

## Collision resolution

The pre-resolution inspection found two duplicate component names and seven
duplicate public standard-module procedure names. Resolution was limited to
the package boundary:

- the three application-event classes now have role-qualified component names;
- standalone `Auto_Open` wrappers remain in the legacy projects but are
  excluded from the shadow, where `modOperationsInit.Auto_Open` is the sole
  startup;
- colliding public helpers were made private where no external caller existed;
- role action/search entry points were role-qualified where callers crossed a
  component boundary; and
- the unreferenced empty `ufDynItemSearchTemplate` is excluded from the shadow,
  while each role-specific search form remains present.

The final deterministic report scans 34 shadow components, 304 public
standard-module procedures, and 8 legacy Ribbon callback names with zero
component, public-procedure, Ribbon callback, or unresolved collision groups.

## GREEN and regression evidence

| Evidence | Result |
|---|---|
| `tests/tooling/Test-Slice6ShadowOperations.ps1` | 10 passed, 0 failed |
| `tools/validate-operations-shadow.ps1` | 13 passed, 0 failed |
| `tools/validate_phase6_packaged_xlams.ps1` | 59 passed, 0 failed |
| `tests/tooling/Test-Slice0ToolContracts.ps1 -Mode All` | 62 passed, 0 failed |
| `tests/tooling/Test-Slice3Baseline.ps1` | 19 passed, 0 failed |
| Shadow collision groups | 0 component, 0 public procedure, 0 Ribbon, 0 unresolved |
| Shadow form initialization | Receiving, Production, and Shipping all PASS |
| Shadow startup/reference/load order | PASS |
| Shadow Ribbon registration | absent, PASS |
| Active legacy role package hashes changed by shadow build | 0 |
| `deploy/current/invSys.Operations.xlam` published | no |

The Slice 5 locks now stand at 8/13: the shadow package definition is GREEN,
while Production/Shipping modeless launch, the Receiving Purchasing stub, the
tabbed Shipping shell, and the real unified Operations Ribbon remain RED for
their planned later slices.

## Static maintenance evidence

The maintenance baseline was regenerated and reproduced deterministically.

| Metric | Slice 5 | Slice 6 | Delta |
|---|---:|---:|---:|
| Packages | 7 | 8 | +1 shadow package |
| Components | 152 | 156 | +4 |
| Procedures | 4,462 | 4,471 | +9 |
| Dynamic roots | 768 | 769 | +1 |
| Scanner warnings | 42 | 42 | 0 |
| Maintenance candidates | 1,043 | 1,047 | +4 |
| Duplicate-body candidates | 193 | 193 | 0 |
| Unresolved dynamic calls | 59 | 59 | 0 |
| Literal `Application.Run` calls | 14 | 14 | 0 |

Explicit Slice 6 growth exception: the four net-new VBA components are the
three six-line standalone startup wrappers required to preserve the active
legacy packages and the 23-line registration-only Operations startup. The
remaining procedure growth is limited to Receiving/Shipping form-initialization
test seams. No duplicate-body, warning, or dynamic-call metric grew.
