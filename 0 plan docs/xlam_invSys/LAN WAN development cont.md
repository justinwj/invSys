# LAN + WAN Development — Correct Path
**Project:** invSys  
**Spec:** `invSys-Design-v4.7.md` (authoritative)  
**Status:** WAN does NOT work. This document replaces the slice scaffold and defines the correct development path to reach: two independent warehouses (WH1 + WH2), each manageable over WAN via SharePoint.

---

## What "WAN working" means

WAN is not a separate connectivity layer you code. WAN IS SharePoint.

The spec (D2) defines it explicitly:

> Each warehouse publishes outbox workbooks and periodic snapshot workbooks to a **SharePoint team document library** when online. HQ aggregates events and produces a global snapshot workbook.

So "two warehouses over WAN" means:

1. WH1 runs on one PC (e.g. X1-Pro-Ai), processes events, produces `WH1.Outbox.Events.xlsb` and `WH1.invSys.Snapshot.Inventory.xlsb`
2. WH2 runs on a second PC (e.g. Arctic-Raptor or ASUS Zenbook), processes events, produces its own outbox + snapshot
3. Both warehouses copy their outbox + snapshot to SharePoint `/invSys/Events/` and `/invSys/Snapshots/`
4. HQ Aggregator (`invSys.HQ.Aggregator.xlsm`) reads both warehouse snapshots from SharePoint → produces `invSys.Global.InventorySnapshot.xlsb`
5. Any PC can open the global snapshot and see advisory cross-warehouse totals

That's it. There is no socket, no API, no real-time sync. The "WAN" is the SharePoint file copy.

---

## Why the previous path was wrong

The LAN WAN development cont.md file was a slice scaffold (distribution contract, tester bundle writer, etc.) focused on tester distribution mechanics — not on the actual WAN data flow. It did not address:

- WH2 runtime provisioning (config, auth, data, snapshot workbooks)
- The SharePoint publish path from each warehouse processor
- The HQ Aggregator's SharePoint read path
- The two-warehouse convergence test

LAN use-within-one-warehouse was also never fully proven (snapshot refresh on a second LAN station was broken). Jumping to WAN before LAN works is the wrong order.

---

## Correct development sequence

### Gate 1 — One-account local (WH1 only) — PREREQUISITE

This must be solid before anything else. On the machine that hosts WH1:

- [ ] `WH1.invSys.Config.xlsb` provisioned with correct `PathDataRoot`, `PathSharePointRoot`
- [ ] `WH1.invSys.Auth.xlsb` provisioned with at least one user per role
- [ ] `WH1.invSys.Data.Inventory.xlsb` exists and schema self-heals on open
- [ ] Receiving → post → processor run → `tblInventoryLog` row → snapshot rebuild → operator `invSys` table refreshes with nonzero row count
- [ ] Shipping and Production same proving path
- [ ] Admin Admin XLAM can break lock, view poison queue, reissue

**Acceptance:** End-to-end role → event → inventory → snapshot → read-model works without Immediate Window intervention.

---

### Gate 2 — LAN (WH1, multiple stations) — PREREQUISITE FOR WAN

A second PC (Arctic-Raptor or Zenbook) as Station S2 on WH1:

- [ ] `net use` authenticated SMB session to `\\WH1-host\invSysWH1` confirmed with `net use` output showing mapped Local letter
- [ ] VBA `FileSystemObject.FileExists` resolves `WH1.invSys.Snapshot.Inventory.xlsb` over that path
- [ ] Excel `Workbooks.Open` can open the snapshot workbook (not just PowerShell `Test-Path`)
- [ ] `modConfig.LoadConfig("WH1", "S2")` returns `True` from the station
- [ ] `modAuth.CanPerform("RECEIVE_POST", currentUser, "WH1", "S2")` returns `True`
- [ ] S2 operator workbook `InventoryManagement!invSys` refreshes to nonzero row count
- [ ] S1 posts event → processor runs on WH1 host → both S1 and S2 refresh to same totals

**Acceptance:** LAN role-usage acceptance standard (v4.7 Phase 6 section) all 7 points satisfied.

---

### Gate 3 — WH2 provisioning (independent warehouse, local only first)

Before any WAN work, WH2 must run independently on its own PC:

- [ ] Create `WH2.invSys.Config.xlsb` — `WarehouseId = WH2`, own `PathDataRoot`, own `PathSharePointRoot` pointing to same SharePoint library root
- [ ] Create `WH2.invSys.Auth.xlsb` with WH2 users
- [ ] Create `WH2.invSys.Data.Inventory.xlsb`
- [ ] WH2 processor runs and produces `WH2.invSys.Snapshot.Inventory.xlsb` and `WH2.Outbox.Events.xlsb`
- [ ] WH2 operator workbook refreshes from WH2 snapshot

**Note:** WH2 runs completely independently of WH1 at this stage. No cross-warehouse path yet.

---

### Gate 4 — SharePoint publish path (WAN layer)

This is the actual WAN implementation. For each warehouse:

**VBA code needed in processor / Admin XLAM:**

```vba
' After GenerateWarehouseSnapshot completes:
Sub PublishToSharePoint(whId As String)
    Dim localSnapshot As String
    Dim localOutbox As String
    Dim spSnapPath As String
    Dim spEventsPath As String

    localSnapshot = PathDataRoot & whId & ".invSys.Snapshot.Inventory.xlsb"
    localOutbox   = PathDataRoot & whId & ".Outbox.Events.xlsb"

    ' PathSharePointRoot from Config, e.g.:
    ' C:\Users\justinwj\[org]\invSys - Documents\
    ' (local OneDrive sync path for SharePoint library)
    spSnapPath   = PathSharePointRoot & "Snapshots\" & whId & ".invSys.Snapshot.Inventory.xlsb"
    spEventsPath = PathSharePointRoot & "Events\" & whId & ".Outbox.Events.xlsb"

    ' FileCopy is atomic enough for this use case (single-file overwrite)
    ' If SharePoint path unreachable, log warning and continue — do NOT block local operation
    On Error GoTo PublishFailed
    FileCopy localSnapshot, spSnapPath
    FileCopy localOutbox, spEventsPath
    LogInfo "Published " & whId & " snapshot and outbox to SharePoint"
    Exit Sub
PublishFailed:
    LogWarning "SharePoint publish failed for " & whId & ": " & Err.Description
End Sub
```

**Requirements:**
- `PathSharePointRoot` in Config points to the **local OneDrive/SharePoint sync folder** (not a URL). On Windows this is typically:
  ```
  C:\Users\justinwj\[OrgName]\invSys - Documents\
  ```
  SharePoint sync client handles the upload automatically.
- Publish must be non-blocking — SharePoint unavailability must never stop local processor operation (per v4.7 invariant).
- WH1 and WH2 each publish to the SAME SharePoint library root, into their own warehouse-named files.

**Proving steps:**
- [ ] WH1 processor run produces snapshot → `FileCopy` to SharePoint sync folder → OneDrive syncs it → visible in SharePoint web UI
- [ ] WH2 processor run does the same for WH2 files
- [ ] SharePoint `/invSys/Snapshots/` contains both `WH1.invSys.Snapshot.Inventory.xlsb` and `WH2.invSys.Snapshot.Inventory.xlsb`
- [ ] SharePoint `/invSys/Events/` contains both outbox files

---

### Gate 5 — HQ Aggregator reads from SharePoint

The HQ Aggregator (`invSys.HQ.Aggregator.xlsm`) reads warehouse snapshots from SharePoint:

**Required behavior:**
1. Copy each `WHx.invSys.Snapshot.Inventory.xlsb` from SharePoint sync folder to a local temp folder (prevents reading a partially-synced file)
2. Open the temp copy, read `tblSkuBalance` (or inventory log)
3. Append rows to the global snapshot with `WarehouseId` stamped
4. Save `invSys.Global.InventorySnapshot.xlsb` to SharePoint `/invSys/Global/`

**Proving steps:**
- [ ] WH1 receives 10 × SKU-001, WH2 receives 5 × SKU-001
- [ ] Both processor runs complete, both snapshots published to SharePoint
- [ ] HQ Aggregator runs → global snapshot shows: WH1 SKU-001 = 10, WH2 SKU-001 = 5 (NOT combined; per-warehouse rows per spec)
- [ ] Global snapshot `SourceType` column shows `SHAREPOINT`
- [ ] Global snapshot opens on any connected PC

---

### Gate 6 — Delayed sync / stale artifact resilience

- [ ] Disconnect WH2 from internet mid-session → WH2 continues processing locally
- [ ] WH2 reconnects → SharePoint sync client resumes → next processor publish succeeds
- [ ] HQ Aggregator run against stale WH2 snapshot shows `IsStale = True` for WH2 rows
- [ ] WH2 operator workbook shows `IsStale = True` when snapshot is from previous session

---

## SharePoint folder structure reminder (from spec)

```
SharePoint: /invSys
├── Addins/Current/          ← XLAMs for distribution
├── Events/
│   ├── WH1.Outbox.Events.xlsb
│   └── WH2.Outbox.Events.xlsb
├── Snapshots/
│   ├── WH1.invSys.Snapshot.Inventory.xlsb
│   └── WH2.invSys.Snapshot.Inventory.xlsb
├── Global/
│   └── invSys.Global.InventorySnapshot.xlsb
├── Config/
├── Auth/
├── Backups/
└── Docs/
```

WH1 and WH2 each write to their own named files. HQ reads all warehouse snapshots from `/Snapshots/`.

---

## Key invariants (never violate)

1. **SharePoint unavailability never blocks local processor operation.** Each warehouse runs offline-first. Publishing is best-effort, logged if it fails.
2. **Global snapshot is advisory only.** WH1's `WH1.invSys.Data.Inventory.xlsb` is the only authoritative store for WH1. Same for WH2.
3. **No cross-warehouse writes.** WH1 processor never touches WH2 workbooks, and vice versa.
4. **WH2 is provisioned independently.** It is not a copy of WH1 runtime — it has its own Config, Auth, Data, Snapshot, Outbox with `WH2` prefix.
5. **HQ Aggregator copies before reading.** Never open a SharePoint-synced file directly — copy to temp first.
6. **Gates are sequential.** WAN (Gate 4+) cannot be claimed if LAN (Gate 2) is not verified.

---

## What needs to be built / fixed

| Item | Status | Notes |
|---|---|---|
| WH1 one-account single-station proving | Unknown — re-verify | Snapshot refresh on operator workbook must be confirmed |
| WH1 LAN multi-station (S2 on second PC) | Not proven | SMB auth + Excel open both required |
| WH2 provisioning (Config, Auth, Data, Snapshot, Outbox) | Not started | Independent of WH1 |
| `PublishToSharePoint` VBA in processor/Admin | Not implemented | FileCopy to local SP sync folder |
| HQ Aggregator reads from SharePoint sync path | Partially implemented | Needs temp-copy safety + WH2 support |
| HQ Aggregator global snapshot per-warehouse rows | Implemented per Phase 5 test | Re-verify against real two-warehouse run |
| Stale snapshot detection on WH2 read model | Not proven | `IsStale`, `LastRefreshUTC` contract |
| Task Scheduler for HQ Aggregator | Not done | Phase 5 open item |
