# Tester Distribution Contract

Status: Authoritative
Scope: WH1 tester onboarding and Confirm Writes proving
Acceptance target: All subsequent LAN/WAN continuation slices

## SharePoint Required Paths

Relative to `PathSharePointRoot`:

```text
Addins/invSys.Core.xlam
Addins/invSys.Inventory.Domain.xlam
Addins/invSys.Receiving.xlam
Addins/invSys.Admin.xlam
TesterPackage/WH1/WH1.TesterBundle.zip
TesterPackage/WH1/WH1.TesterReadme.md
```

## Local Machine Required State After First-Run Bootstrap

```text
C:\invSys\WH1\
C:\invSys\WH1\config\
C:\invSys\WH1\auth\
C:\invSys\WH1\inbox\
C:\invSys\WH1\outbox\
C:\invSys\WH1\snapshots\
C:\invSys\WH1\WH1.invSys.Data.Inventory.xlsb
C:\invSys\WH1\WH1.Receiving.Operator.xlsm
```

## Tester Auth Required State

```text
UserId: tester-defined
WarehouseId: WH1
StationId: R1
Capabilities: RECEIVE_POST, RECEIVE_VIEW, READMODEL_REFRESH
Status: ACTIVE
```

## Confirm Writes Seeded Scenario

```text
WarehouseId: WH1
StationId: R1
SKU: TEST-SKU-001
Initial QtyOnHand: 100
Test action: receive 10 units of TEST-SKU-001
Expected postcondition: QtyOnHand = 110 in read-model after processor runs
```

## Validation Rules

```text
Rule: SetupTesterStation completion is valid only if every SharePoint required path exists.
Rule: SetupTesterStation completion is valid only if every local machine required path exists.
Rule: WH1.Receiving.Operator.xlsm must exist at the exact local path defined above.
Rule: Tester auth must be scoped to WarehouseId=WH1 and StationId=R1.
Rule: Tester auth must include RECEIVE_POST, RECEIVE_VIEW, and READMODEL_REFRESH.
Rule: Tester auth status must be ACTIVE.
Rule: Confirm Writes proving is valid only against TEST-SKU-001 with initial QtyOnHand=100 and expected QtyOnHand=110.
```
