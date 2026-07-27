# Slice 0 tool-contract results

**Slice:** 0 — Tool contracts, schemas, and synthetic fixtures

**Recorded:** 2026-07-27 12:07:08 -07:00

**Runtime workbook required:** No

**Operational workbook opened or changed:** No

## Contract validation

Command:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\tooling\Test-Slice0ToolContracts.ps1 -Mode Contracts
```

Result:

```text
RESULT passed=36 failed=0
```

The passing contract set proves:

- four distinct versioned JSON contracts are parseable and expose their
  required static, maintenance, runtime, and dynamic-root sections;
- the static fixtures contain direct, literal cross-XLAM, unresolved dynamic,
  event, duplicate-body, and unreachable cases;
- the runtime fixtures distinguish managed `System_Key`/`Condition` headers
  from preserved `Custom_Color`;
- the negative fixture contains retired `ROW` solely so the tools must flag it;
- all three synthetic credential classes are present in input and absent from
  expected evidence; and
- expected Markdown fixtures are byte-stable.

## Meaningful RED

Command:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\tooling\Test-Slice0ToolContracts.ps1 -Mode All
```

Result:

```text
RESULT passed=36 failed=2
ToolA.EntryPoint: tools/inventory-vba-surface.ps1 is absent; this is the expected Slice 0 RED.
ToolB.EntryPoint: tools/export-invsys-runtime-state.ps1 is absent; this is the expected Slice 0 RED.
```

The two failures are the intended behavioral gaps. The harness, schemas, and
fixtures are green; Tool A and Tool B do not yet implement their frozen command
and output contracts. Slice 1 must make `-Mode Static` green. Slice 2 must make
`-Mode Runtime` green. Neither slice may weaken the assertions to claim GREEN.

## Safety

The tests read only committed text/JSON fixtures and create disposable output
only under an `invsys-slice0-*` directory inside the Windows temporary
directory. Cleanup verifies both the temporary parent and the invSys-specific
name prefix before recursive removal.
