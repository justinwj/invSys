# Slice 7 Production Session and Completion Service

Date: 2026-07-27

## Scope

Slice 7 replaces worksheet/control authority for an active Production run with
a typed session and completion result:

- consumed inventory is captured by immutable `System_Key`;
- the output `System_Key` is allocated once at the Production creation
  boundary;
- consume and complete events have distinct, stable event IDs;
- readiness for the next batch requires verified processor application and
  read-model refresh;
- partial application records whether compensation is required; and
- hidden workbook-name metadata preserves the session across save, close,
  reopen, and reload without making a worksheet table authoritative.

Production completion queues canonical events and invokes the processor and
operator refresh boundary. It does not mutate canonical inventory directly.

## D13 trace

The focused session tests were written and run before the service existed.
The initial meaningful RED was 0/7 because
`modProductionCompletionService.ProductionSessionContractProbe` was absent.

The first workbook restart test was also observed RED: storing the serialized
session in one defined-name formula exceeded Excel's formula length limit.
Chunked hidden names then made the real save/close/reopen/load path GREEN.

The packaged two-batch form-action target supplied two further meaningful RED
states:

- the action remained at zero-percent allocation because its real run-location
  selector was not populated; and
- after reaching completion, synchronization attempted `ListIndex = 0` on an
  empty peer ComboBox and raised run-time error 380.

The final unchanged action returned
`OK|Batches=2|BoundWorkbook=<Production operator workbook>`.

## GREEN and regression evidence

| Evidence | Result |
|---|---|
| Focused tests 243-251 | 9 passed, 0 failed |
| Actual workbook save/close/reopen session test | PASS |
| System_Key CheckIn staging without canonical mutation | PASS |
| Packaged Production direct CheckIn/complete/process/log workflow | PASS |
| Packaged two-consecutive-batch operator action | PASS |
| `tools/validate_phase6_packaged_xlams.ps1` | 59 passed, 0 failed |
| `tools/validate-operations-shadow.ps1` | 13 passed, 0 failed |
| `tests/tooling/Test-Slice0ToolContracts.ps1 -Mode All` | 62 passed, 0 failed |
| `tests/tooling/Test-Slice3Baseline.ps1` | 19 passed, 0 failed |
| Shadow collision groups | 0 unresolved |

The broader live role validator improved from 26/45 before the Slice 7 work to
34/45. Its remaining failures are outside the Slice 7 Production-session
contract, including Receiving, Shipping/Boxing, and their resulting aggregate
balance expectations. The raw live report contains machine/runtime paths and
is intentionally not release evidence.

## Static maintenance evidence

| Metric | Slice 6 | Slice 7 | Delta |
|---|---:|---:|---:|
| Packages | 8 | 8 | 0 |
| Components | 156 | 159 | +3 |
| Procedures | 4,471 | 4,569 | +98 |
| Maintenance candidates | 1,047 | 1,048 | +1 |
| Duplicate-body candidates | 193 | 193 | 0 |
| Unresolved dynamic calls | 59 | 59 | 0 |
| Literal `Application.Run` calls | 14 | 14 | 0 |

The three new components are the typed run session, structured completion
result, and Production completion service required by this slice. The single
net maintenance candidate is review-only; duplicate-body and dynamic-call
ratchets did not grow.

## Compile-blocker repairs

During packaged validation, stale same-project helper calls were corrected to
their role-qualified procedures:

- Receiving calls `ShowReceivingDynamicItemSearch`;
- Shipping calls `ShowShippingDynamicItemSearch`.

The harness imports the current role-qualified Receiving application-event
class. These repairs restored compile evidence without changing a runtime
contract.
