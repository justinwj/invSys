# LAN additions

## Purpose

This addendum operationalizes Release 1 LAN setup so role usage is dependable across multiple stations without relying on ad hoc Immediate Window work, accidental local runtime creation, or ambiguous Windows share state.

This document does not replace the main design spec. It turns the existing LAN intent into a concrete end-user operating model, bootstrap sequence, validation sequence, and troubleshooting baseline.

***

## Scope

This document covers one warehouse, multiple LAN stations, shared local-network runtime artifacts, and role workbooks using snapshot-fed `invSys` read models.

It applies to:
- Receiving
- Shipping
- Production
- Admin, where relevant for orchestration and auth maintenance

It does not define WAN publication or HQ aggregation. Those remain separate proving layers.

***

## Problem This Addendum Solves

The architecture already says:
- each station writes to its own inbox workbook
- one warehouse runtime is authoritative
- operator `invSys` is a snapshot-fed read model

What was not sufficiently operationalized was the end-user path needed to make that true on a second PC:
- which files must be shared vs local
- which credentials must be used for SMB access
- how station auth is provisioned
- how the managed inventory list becomes available on the second station
- how to distinguish a runtime problem from an Excel/share/session problem

This addendum closes that gap.

***

## Authoritative LAN Model

### Shared warehouse runtime

One warehouse host owns the authoritative warehouse runtime path.

Example:
```text
X1-Pro-Ai
C:\invSys\WH1
\\X1-Pro-Ai\invSysWH1
\\192.168.1.5\invSysWH1
```

This shared warehouse runtime contains:
- `WH1.invSys.Config.xlsb`
- `WH1.invSys.Auth.xlsb`
- `WH1.invSys.Data.Inventory.xlsb`
- `WH1.invSys.Snapshot.Inventory.xlsb`
- `WH1.Outbox.Events.xlsb`
- any other warehouse-authoritative runtime artifacts

These files are warehouse-owned, not station-owned.

### Station-local operator context

Each station owns:
- its local role operator workbook
- its own station inbox workbook
- optionally its own local config copy used for operator/runtime bootstrap

Example for Arctic-Raptor `S2`:
```text
Operator workbook:
C:\Users\justinwj\Documents\WH1_S2_Receiving_Operator.xlsb

Local station config copy:
C:\invSys\WH1\WH1.invSys.Config.xlsb

Station inbox root:
\\192.168.1.3\invSysStationS2

Station inbox workbook:
\\192.168.1.3\invSysStationS2\invSys.Inbox.Receiving.S2.xlsb
```

### Source-of-truth rule

The source of truth remains the canonical warehouse inventory workbook:
```text
WH1.invSys.Data.Inventory.xlsb
```

The snapshot workbook is not authoritative.

The operator `invSys` table is not authoritative.

The outbox is not authoritative.

***

## Required Config Semantics

### `tblWarehouseConfig`

`PathDataRoot` must resolve to a path Excel can actually open from the station machine.

Release 1 requirement:
- it is not enough for PowerShell `Test-Path` to succeed
- Excel/VBA must also be able to open the snapshot workbook through that path

If Excel cannot reliably open a UNC path from the station, a mapped drive letter may be used on that station.

Example:
```text
PathDataRoot = W:\
```

Only the station-local config copy should be changed to a station-specific mapped-drive view if needed.

The shared warehouse config should continue to describe the warehouse runtime path from the warehouse perspective.

### `tblStationConfig`

Each station row must contain:
- `StationId`
- `WarehouseId`
- `StationName`
- `PathInboxRoot`
- `RoleDefault`

Example:
```text
StationId     = S2
WarehouseId   = WH1
StationName   = ARCTIC-RAPTOR
PathInboxRoot = \\192.168.1.3\invSysStationS2\
RoleDefault   = RECEIVE
```

### LAN rule

`PathDataRoot` and `PathInboxRoot` serve different purposes:
- `PathDataRoot` points to shared warehouse runtime artifacts
- `PathInboxRoot` points to the station’s own inbox location

They must not be conflated.

***

## Required Auth Semantics

LAN station setup is not complete until the station user has explicit auth rows in the shared warehouse auth workbook.

### Why this matters

A second station can have:
- valid snapshot read access
- a populated `invSys` table
- a working inbox workbook

and still fail `Confirm Writes` if the current Windows user lacks `RECEIVE_POST`, `SHIP_POST`, or `PROD_POST`.

### Required shared auth provisioning

In `WH1.invSys.Auth.xlsb`:

`tblUsers` must include the station user.

Example:
```text
UserId      = justinwj
DisplayName = justinwj
Status      = Active
```

`tblCapabilities` must include role-appropriate capabilities.

Receiving example:
```text
UserId      = justinwj
Capability  = RECEIVE_POST
WarehouseId = WH1
StationId   = *
Status      = ACTIVE
```

Recommended minimum grants by role:
- Receiving station: `RECEIVE_POST`
- Shipping station: `SHIP_POST`
- Production station: `PROD_POST`
- Processor service: `INBOX_PROCESS`

### R1 operational requirement

Station bootstrap should be considered incomplete until:
- station config exists
- station inbox exists
- station user auth exists

Auth provisioning is therefore part of dependable LAN setup, not an optional afterthought.

***

## Managed Inventory Availability Rule

### Definition

The “managed inventory list” available to a role station is the local operator workbook’s `InventoryManagement!invSys` table after snapshot refresh.

It is not a separate replicated catalog workbook.

It is not populated from local staging tables.

It is not station-private truth.

### Implication

For a second station to have usable managed inventory:
1. the station must load the shared runtime config successfully
2. Excel on that station must be able to open the warehouse snapshot workbook
3. the operator workbook must refresh `InventoryManagement!invSys`
4. the operator workbook must be the active workbook if the active-workbook wrapper macro is used

### Required validation

A station is not considered role-ready until these are true:
```vb
?Application.Run("'invSys.Core.xlam'!modOperatorReadModel.RefreshInventoryReadModelForWorkbook", Workbooks("WH1_S2_Receiving_Operator.xlsb"), "WH1", "LOCAL")
True
```

```vb
?Workbooks("WH1_S2_Receiving_Operator.xlsb").Worksheets("InventoryManagement").ListObjects("invSys").ListRows.Count
```

Row count must be greater than zero for an inventory-populated warehouse.

***

## SMB and Excel Access Requirements

### Critical rule

Windows shell access is not sufficient proof of Excel access.

The following all must be distinguished:
- PowerShell `Test-Path`
- File Explorer access
- VBA `FileSystemObject.FileExists`
- Excel `Workbooks.Open`

A station can pass shell checks and still fail Excel/VBA file opens.

### Required share authentication

SMB access must be authenticated with an explicit warehouse share account or another approved account with read/write permission.

Example:
```powershell
net use \\192.168.1.5\invSysWH1 /user:X1-PRO-AI\invsyslan * /persistent:yes
```

Do not rely on accidental guest/anonymous or wrong-context sessions.

### Validation ladder

#### 1. Shell-level
```powershell
Get-ChildItem "\\192.168.1.5\invSysWH1"
```

#### 2. Excel/VBA file visibility
```vb
?CreateObject("Scripting.FileSystemObject").FileExists("\\192.168.1.5\invSysWH1\WH1.invSys.Snapshot.Inventory.xlsb")
```

#### 3. Excel workbook open
Excel must be able to open the snapshot workbook without a 1004 open failure.

### Mapped drive fallback

If Excel/VBA cannot reliably open the UNC path, a station-local mapped drive may be used.

Example:
```powershell
net use W: \\192.168.1.5\invSysWH1 /user:X1-PRO-AI\invsyslan * /persistent:yes
```

Then station-local `PathDataRoot` may be:
```text
W:\
```

This is a station-local compatibility workaround, not a change to warehouse authority.

### Important rule

A mapped drive is not real until `net use` shows a `Local` drive letter mapping and File Explorer can browse it.

Shell probes alone are not enough.

***

## Required End-User LAN Bootstrap Sequence

### Warehouse host setup

On the warehouse host:
1. create/maintain the canonical warehouse runtime folder
2. share it over SMB
3. grant the designated LAN account the required read/write access
4. confirm the shared warehouse runtime contains config/auth/inventory/snapshot/outbox files

### Station setup

On each station:
1. install or copy the rebuilt `deploy/current` XLAMs locally
2. ensure access to the shared warehouse runtime via authenticated SMB
3. create/share the station inbox root if the processor must reach it over LAN
4. run station bootstrap to create:
   - local config copy
   - station inbox workbook
   - operator workbook
5. ensure shared auth grants the station user the required role capability
6. verify Excel can open the snapshot path
7. refresh the operator read model and confirm `invSys` row count is nonzero

### Role-ready acceptance criteria

A station is only role-ready when all are true:
- shared runtime reachable from station
- shared auth reachable from station
- station inbox reachable from warehouse processor
- operator workbook exists
- `invSys` refresh succeeds
- `invSys` shows rows
- current user has role capability

***

## Wrapper Macro Activation Rule

`RefreshCurrentWorkbookInventoryReadModel` uses the active workbook context.

If the active workbook is:
- config
- auth
- snapshot
- any non-operator workbook

then the wrapper can correctly report:
```text
invSys table not found.
```

This is not necessarily a read-model failure.

### Rule

For deterministic station operations:
- activate the operator workbook before using the active-workbook wrapper
- or use the workbook-targeted function directly

Preferred deterministic call:
```vb
?Application.Run("'invSys.Core.xlam'!modOperatorReadModel.RefreshInventoryReadModelForWorkbook", Workbooks("WH1_S2_Receiving_Operator.xlsb"), "WH1", "LOCAL")
```

***

## Operator Workflow Dependability Requirements

### Receiving

To call Receiving dependable on LAN:
- item picker must load from populated `InventoryManagement!invSys`
- `Confirm Writes` must enqueue to station inbox
- processor must apply event and rebuild snapshot
- both stations must refresh to converged totals

### Shipping

To call Shipping dependable on LAN:
- shipping staging must remain local
- `invSys` read model must refresh non-destructively
- `SHIP_POST` must be granted to the station user
- shipment events must serialize through warehouse processor

### Production

To call Production dependable on LAN:
- production staging must remain local
- `invSys` read model must refresh non-destructively
- `PROD_POST` must be granted to the station user
- production events must serialize through warehouse processor

***

## Minimum LAN Validation Checklist

### Station health

On each station:
- `modConfig.LoadConfig(warehouse, station)` returns `True`
- `PathDataRoot` resolves to an Excel-openable path
- `PathInboxRoot` resolves to the station inbox location
- `modAuth.LoadAuth(warehouse)` returns `True`
- `modAuth.CanPerform(roleCapability, currentUser, warehouse, station, ...)` returns `True`

### Read-model health

On each station:
- snapshot workbook resolves
- snapshot table resolves
- snapshot row count is nonzero when warehouse has inventory
- `invSys` row count is nonzero

### Write-path health

Per station:
- role post succeeds
- inbox row becomes `NEW`
- processor run marks it `PROCESSED`
- canonical inventory log records the event
- snapshot refresh exposes the change on both stations

### Locking health

Across stations:
- competing process attempts do not corrupt data
- one lane wins cleanly
- retry after release succeeds

***

## Troubleshooting Matrix

### Symptom: `invSys` table visually blank on second station

Check:
- `ListRows.Count`
- direct workbook-targeted refresh
- whether the operator workbook is active

Likely causes:
- snapshot not reachable
- wrapper targeting wrong workbook
- table populated but sheet focus/filters misleading user

### Symptom: `Snapshot workbook not found; operator read model marked stale.`

Check:
- station `PathDataRoot`
- shell access
- Excel/VBA `FileExists`
- Excel workbook open by path

Likely causes:
- unauthenticated SMB session
- mapped drive not real in Windows shell context
- Excel cannot open UNC path even though PowerShell can

### Symptom: `Current user lacks RECEIVE_POST capability.`

Check:
- `tblUsers`
- `tblCapabilities`
- current Windows user id
- whether station auth data was actually provisioned

Likely cause:
- station user exists operationally but was never added to shared auth

### Symptom: `invSys table not found.`

Check:
- which workbook is active
- whether the operator workbook is the current active workbook

Likely cause:
- wrapper macro called while config/auth/snapshot workbook is active

***

## Required Productization Follow-Up

The following are now mandatory for dependable end-user LAN usage:

1. `setup_lan_station.ps1` should optionally provision shared auth rows for the station user.
2. Station bootstrap should include a post-bootstrap validation report that explicitly proves:
   - shell access
   - Excel/VBA file access
   - snapshot open
   - `invSys` row count
3. Role refresh entry points should prefer workbook-targeted refresh or explicitly activate the operator workbook before using active-workbook wrappers.
4. LAN setup docs should state that authenticated SMB access must be proven at both shell and Excel/VBA levels.
5. “Managed inventory available” should be an explicit readiness gate, not an assumed byproduct of sheet creation.

***

## Acceptance Standard for “LAN Role Usage Dependable”

LAN role usage is dependable only when all of the following are true:

1. Multiple stations can open role workbooks against one warehouse runtime.
2. Each station can refresh `invSys` from the warehouse snapshot without local workbook contamination.
3. Each station user has the required auth capability.
4. Each station posts only to its own inbox workbook.
5. Warehouse processor serializes canonical writes and snapshot rebuilds correctly.
6. Two stations converge to the same visible inventory totals after refresh.
7. The above works without Immediate Window intervention beyond diagnostics.

If any of those are false, LAN architecture may be partially proven, but LAN end-user operation is not yet dependable.
