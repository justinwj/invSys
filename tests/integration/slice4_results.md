# Slice 4 — Greenfield warehouse generation and demo inventory seeding

**Recorded:** 2026-07-27

**Normative contract:** D14 clean-start inventory identity

## D13 trace

Meaningful RED was observed before implementation:

- `D14.GeneratedSchemas.ManagedHeadersNoROW` failed because generated Inventory
  Domain, snapshot, and operator tables lacked `System_Key`/`Condition` and
  exposed `ROW`.
- `D14.AdminSeed.RepeatedCreatesUniqueKeysNoMigration` failed because the Admin
  seed action used the legacy migration path and did not create new keyed
  entities on each run.
- Phase 6 test 105 processed a RECEIVE event but the operator read model could
  not identify the new entity because the creation boundary did not allocate a
  `System_Key`.
- Phase 6 test 111 preserved keys and quantities but exposed blank operator
  locations, tracing a lost entity location to snapshot summary normalization.

GREEN evidence:

| Evidence | Result |
|---|---|
| `tools/run_create_warehouse_integration.ps1` | 15 passed, 0 failed |
| Phase 6 tests 94–112 | 19 passed, 0 failed |
| `tools/validate_phase6_packaged_xlams.ps1` | 59 passed, 0 failed |
| `tests/tooling/Test-Slice0ToolContracts.ps1 -Mode All` | 62 passed, 0 failed |
| `tests/tooling/Test-Slice3Baseline.ps1` | 19 passed, 0 failed |
| `tools/build-xlam.ps1 -Apply` | 7 current XLAMs built |
| Evidence path scan | no machine-specific absolute path in committed runtime evidence |

The clean-start test proves:

- all generated durable entities have unique nonblank keys;
- seeded `Condition` is `GOOD`;
- Inventory Domain → snapshot → operator refresh preserves the exact keys;
- two repeated Admin seed actions add six new keys without collision or
  migration events;
- an additional operator column survives refresh and save/reopen by key; and
- generated managed tables contain no `ROW` header.

## Static maintenance evidence

The maintenance baseline was regenerated after implementation and then reproduced
deterministically. Compared with the Slice 3 baseline:

| Metric | Before | Slice 4 | Delta |
|---|---:|---:|---:|
| Components | 151 | 152 | +1 |
| Procedures | 4,441 | 4,455 | +14 |
| Dynamic roots | 768 | 768 | 0 |
| Scanner warnings | 42 | 42 | 0 |
| Maintenance candidates | 1,038 | 1,041 | +3 |
| Oversized modules | 25 | 25 | 0 |
| Unresolved dynamic calls | 0 | 0 | 0 |
| Same-project `Application.Run` candidates | 0 | 0 | 0 |
| Duplicate bodies | 0 | 0 | 0 |

Explicit Slice 4 growth exception:

| Existing module | Line delta | Required contract work |
|---|---:|---|
| `modAdmin` | +5 | packaged Admin seed delegates to the headless seed service |
| `modOperatorReadModel` | +15 | key-addressed projection and custom-column preservation |
| `modProcessor` | +19 | keyed inbox envelope and `ROW` removal |
| `modRoleEventWriter` | +90 | creation-time key allocation, serialization, and normalized GUIDs |
| `modWarehouseBootstrap` | +1 | clean-start seed event |
| `modWarehouseSync` | +40 | entity-keyed snapshot and location preservation |
| `modInventoryApply` | +134 | strict create/receive key validation and entity projection |

The new `modAdminInventorySeed` service is 134 lines, below the 1,000-line new
module limit. The existing-module growth is accepted only for the D14 foundation;
the plan’s later service-extraction slices remain responsible for reducing the
large role/controller modules. No growth is allowed by default in the regenerated
ratchets.
