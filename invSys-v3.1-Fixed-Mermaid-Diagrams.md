# Fixed Mermaid Diagrams for invSys v3.1

## System topology (Python HQ Aggregator)

```mermaid
flowchart TB
  subgraph Warehouse1["Warehouse 1 - LAN-First"]
    W1Stations[Receiving/Shipping/Production Stations]
    W1Inbox["Station Inbox Workbooks<br/>invSys.Inbox.*.xlsb"]
    W1Proc["Processor Station<br/>Core.Processor - VBA"]
    W1Auth[WH1.invSys.Auth.xlsb]
    W1Inv[WH1.invSys.Data.Inventory.xlsb]
    W1Des[WH1.invSys.Data.Designs.xlsb]
    W1Out[WH1.Outbox.Events.xlsb]
    W1Snap[WH1.invSys.Snapshot.Inventory.xlsb]

    W1Stations -- "Append events<br/>NEW rows" --> W1Inbox
    W1Inbox -- "Batch process<br/>VBA Processor" --> W1Proc
    W1Proc -- Read --> W1Auth
    W1Proc -- Write --> W1Inv
    W1Proc -- Write --> W1Des
    W1Proc -- "Publish events" --> W1Out
    W1Proc -- "Generate<br/>warehouse snapshot" --> W1Snap
  end

  subgraph Warehouse2["Warehouse 2 - LAN-First"]
    W2Stations[Receiving/Shipping/Production Stations]
    W2Inbox[Station Inbox Workbooks]
    W2Proc["Processor Station<br/>Core.Processor - VBA"]
    W2Inv[WH2.invSys.Data.Inventory.xlsb]
    W2Out[WH2.Outbox.Events.xlsb]

    W2Stations -- "Append events" --> W2Inbox
    W2Inbox -- "Batch process" --> W2Proc
    W2Proc -- Write --> W2Inv
    W2Proc -- "Publish events" --> W2Out
  end

  subgraph SharePoint["SharePoint Online<br/>via OneDrive sync"]
    SPAddins[Addins - XLAM distribution]
    SPEvents[Events - Outbox uploads]
    SPGlobal[Global Snapshots]
    SPConfig[Config workbooks]
    SPAuth[Auth workbooks]
  end

  subgraph HQ["HQ Aggregation - Python"]
    HQTask[Windows Task Scheduler]
    HQPy["invSys_hq_aggregator.py<br/>Python"]
    HQGlobal[Global.InventorySnapshot.xlsb]
  end

  W1Out -. "Sync when online" .-> SPEvents
  W2Out -. "Sync when online" .-> SPEvents
  W1Snap -. "Sync when online" .-> SPGlobal

  SPEvents -- "Read .xlsb via pyxlsb<br/>no Excel COM" --> HQPy
  HQTask --> HQPy
  HQPy -- "Aggregate snapshots<br/>across warehouses" --> HQGlobal
  HQGlobal -. "Sync via OneDrive" .-> SPGlobal
  SPGlobal -. "Read-only" .-> RemoteBI[Remote BI / Sales]

  style W1Proc fill:#424242,stroke:#1b1b1b,color:#fff
  style W2Proc fill:#424242,stroke:#1b1b1b,color:#fff
  style HQPy fill:#424242,stroke:#1b1b1b,color:#fff
  style W1Inv fill:#2e7d32,stroke:#1b5e20,color:#fff
  style W1Des fill:#2e7d32,stroke:#1b5e20,color:#fff
  style W2Inv fill:#2e7d32,stroke:#1b5e20,color:#fff
  style W1Inbox fill:#ef6c00,stroke:#e65100,color:#fff
  style W2Inbox fill:#ef6c00,stroke:#e65100,color:#fff
  style W1Out fill:#ef6c00,stroke:#e65100,color:#fff
  style W2Out fill:#ef6c00,stroke:#e65100,color:#fff
```

---

## SharePoint folder structure

```mermaid
flowchart TB
  ROOT["SharePoint: /invSys"]

  ROOT --> ADDINS[Addins]
  ROOT --> EVENTS[Events]
  ROOT --> GLOBAL[Global]
  ROOT --> CONFIG[Config]
  ROOT --> AUTH[Auth]
  ROOT --> DOCS[Docs]
  ROOT --> ARCHIVE[Archive]

  ADDINS --> ADDINSCURRENT[Current]
  ADDINS --> ADDINSARCHIVE[Archive]

  ADDINSCURRENT --> XLAMCORE[invSys.Core.xlam]
  ADDINSCURRENT --> XLAMINV[invSys.Inventory.Domain.xlam]
  ADDINSCURRENT --> XLAMDES[invSys.Designs.Domain.xlam]
  ADDINSCURRENT --> XLAMRECV[invSys.Receiving.xlam]
  ADDINSCURRENT --> XLAMSHIP[invSys.Shipping.xlam]
  ADDINSCURRENT --> XLAMPROD[invSys.Production.xlam]
  ADDINSCURRENT --> XLAMADMIN[invSys.Admin.xlam]

  ADDINSARCHIVE --> V11[v1.1]
  ADDINSARCHIVE --> V12[v1.2]

  EVENTS --> EWH1[WH1.Outbox.Events.xlsb]
  EVENTS --> EWH2[WH2.Outbox.Events.xlsb]

  GLOBAL --> GINV[Global.InventorySnapshot.xlsb]
  GLOBAL --> GDES[Global.DesignsSnapshot.xlsb]

  CONFIG --> CWH1[WH1.invSys.Config.xlsb]
  CONFIG --> CWH2[WH2.invSys.Config.xlsb]

  AUTH --> AWH1[WH1.invSys.Auth.xlsb]
  AUTH --> AWH2[WH2.invSys.Auth.xlsb]

  style XLAMCORE fill:#1f78b4,stroke:#0b4f6c,color:#fff
  style XLAMINV fill:#00897b,stroke:#00695c,color:#fff
  style XLAMDES fill:#00897b,stroke:#00695c,color:#fff
  style XLAMRECV fill:#6a1b9a,stroke:#4a148c,color:#fff
  style XLAMSHIP fill:#6a1b9a,stroke:#4a148c,color:#fff
  style XLAMPROD fill:#6a1b9a,stroke:#4a148c,color:#fff
  style XLAMADMIN fill:#6a1b9a,stroke:#4a148c,color:#fff
  style EWH1 fill:#ef6c00,stroke:#e65100,color:#fff
  style EWH2 fill:#ef6c00,stroke:#e65100,color:#fff
  style GINV fill:#2e7d32,stroke:#1b5e20,color:#fff
  style GDES fill:#2e7d32,stroke:#1b5e20,color:#fff
  style CWH1 fill:#fbc02d,stroke:#f9a825,color:#000
  style CWH2 fill:#fbc02d,stroke:#f9a825,color:#000
  style AWH1 fill:#616161,stroke:#424242,color:#fff
  style AWH2 fill:#616161,stroke:#424242,color:#fff
```

---

## Repository layout

```mermaid
flowchart TB
  ROOT["invSys - repo root"]

  ROOT --> SRC[src]
  ROOT --> DATA[data]
  ROOT --> ASSETS[assets]
  ROOT --> DEPLOY[deploy]
  ROOT --> TOOLS[tools]
  ROOT --> DOCS[docs]
  ROOT --> TESTS[tests]

  SRC --> CORE[Core]
  SRC --> INVDOM[InventoryDomain]
  SRC --> DESDOM[DesignsDomain]
  SRC --> RECV[Receiving]
  SRC --> SHIP[Shipping]
  SRC --> PROD[Production]
  SRC --> ADMIN[Admin]

  CORE --> COREM[Modules]
  CORE --> COREC[ClassModules]
  CORE --> CORER[Ribbon]
  CORE --> CORECFG[Config]

  INVDOM --> INVDOMM[Modules]
  INVDOM --> INVDOMC[ClassModules]
  INVDOM --> INVDOMSCHEMA[Schema]

  RECV --> RECVM[Modules]
  RECV --> RECVF[Forms]
  RECV --> RECVR[Ribbon]

  DOCS --> DARCH[architecture]
  DOCS --> DWORK[workflows]
  DOCS --> DROLES[roles-permissions]
  DOCS --> DIMPL[implementation]

  TOOLS --> TPY["Python: hq_aggregator / backup / validation"]
  TOOLS --> TEXPORT[export-vba.ps1]
  TOOLS --> TBUILD[build-xlam.ps1]
  TOOLS --> TSYNC[sync-forms.ps1]
  TOOLS --> TCRYPTO[CryptoHelper.bas]

  TESTS --> TUNIT[unit]
  TESTS --> TINTEGRATION[integration]
  TESTS --> THARNESS[TestHarness.xlsm]

  style CORE fill:#1f78b4,stroke:#0b4f6c,color:#fff
  style INVDOM fill:#00897b,stroke:#00695c,color:#fff
  style DESDOM fill:#00897b,stroke:#00695c,color:#fff
  style RECV fill:#6a1b9a,stroke:#4a148c,color:#fff
  style SHIP fill:#6a1b9a,stroke:#4a148c,color:#fff
  style PROD fill:#6a1b9a,stroke:#4a148c,color:#fff
  style ADMIN fill:#6a1b9a,stroke:#4a148c,color:#fff
```

---

## Component dependency graph

```mermaid
graph TD
  Config["Core.Config - VBA"]
  Auth["Core.Auth - VBA"]
  Lock["Core.LockManager - VBA"]
  Proc["Core.Processor - VBA"]
  InvSchema["InventoryDomain.Schema - VBA"]
  InvApply["InventoryDomain.Apply - VBA"]
  DesSchema["DesignsDomain.Schema - VBA"]
  DesApply["DesignsDomain.Apply - VBA"]
  RecvUI["Receiving.UI - VBA"]
  ShipUI["Shipping.UI - VBA"]
  ProdUI["Production.UI - VBA"]
  AdminUI["Admin.UI - VBA"]
  HQAgg["invSys_hq_aggregator.py - Python"]
  PyBackup["backup_invSys.py - Python"]
  PyValidate["validate_schema.py - Python"]
  PyMonitor["monitor_invSys.py - Python R2+"]
  ItemAPI["item_search_api.py - Python R2"]

  Config --> Auth
  Config --> Lock
  Auth --> RecvUI
  Lock --> Proc
  Auth --> Proc
  InvSchema --> InvApply
  InvApply --> Proc
  Proc --> AdminUI
  Config --> InvSchema
  Config --> DesSchema
  DesSchema --> DesApply
  DesApply --> Proc
  RecvUI --> RecvEvent[Receiving.EventCreator]
  Auth --> ShipUI
  ShipUI --> ShipEvent[Shipping.EventCreator]
  Auth --> ProdUI
  ProdUI --> ProdEvent[Production.EventCreator]

  Proc --> WHOut[WHx.Outbox.Events.xlsb]
  WHOut --> HQAgg

  DATA[.xlsb workbooks] --> HQAgg
  DATA --> PyBackup
  DATA --> PyValidate
  LOGS[Logs / metrics] --> PyMonitor
  PyMonitor --> Email[Email alerts]

  style Config fill:#fbc02d,stroke:#f9a825,color:#000
  style Auth fill:#616161,stroke:#424242,color:#fff
  style Lock fill:#1f78b4,stroke:#0b4f6c,color:#fff
  style Proc fill:#424242,stroke:#1b1b1b,color:#fff
  style InvSchema fill:#00897b,stroke:#00695c,color:#fff
  style InvApply fill:#00897b,stroke:#00695c,color:#fff
  style DesSchema fill:#00897b,stroke:#00695c,color:#fff
  style DesApply fill:#00897b,stroke:#00695c,color:#fff
  style RecvUI fill:#6a1b9a,stroke:#4a148c,color:#fff
  style ShipUI fill:#6a1b9a,stroke:#4a148c,color:#fff
  style ProdUI fill:#6a1b9a,stroke:#4a148c,color:#fff
  style AdminUI fill:#6a1b9a,stroke:#4a148c,color:#fff
  style HQAgg fill:#424242,stroke:#1b1b1b,color:#fff
  style PyBackup fill:#424242,stroke:#1b1b1b,color:#fff
  style PyValidate fill:#424242,stroke:#1b1b1b,color:#fff
  style PyMonitor fill:#424242,stroke:#1b1b1b,color:#fff
```

---

## HQ global snapshot workflow

```mermaid
sequenceDiagram
  participant TaskScheduler as Windows Task Scheduler
  participant HQPy as invSys_hq_aggregator.py
  participant OneDrive as OneDrive Sync Folder
  participant Events as Events Folder
  participant GlobalWB as Global.InventorySnapshot.xlsb

  TaskScheduler->>HQPy: Trigger hourly
  HQPy->>OneDrive: Enumerate Events/*.Outbox.Events.xlsb
  OneDrive-->>HQPy: WH1.Outbox.Events.xlsb, WH2.Outbox.Events.xlsb

  loop For each warehouse outbox
    HQPy->>OneDrive: Read tblOutboxEvents via pyxlsb
    HQPy->>HQPy: Aggregate by SKU/Warehouse<br/>last-write-wins by AppliedAtUTC
  end

  HQPy->>GlobalWB: Overwrite tblGlobalOnHand<br/>write .xlsb via openpyxl
  HQPy->>OneDrive: Save Global.InventorySnapshot.xlsb
  HQPy-->>TaskScheduler: Exit code 0 success
```

---

## Warehouse Processor Batch Application

```mermaid
sequenceDiagram
  participant Admin
  participant AdminUI as Admin UI
  participant Processor as Core.Processor
  participant LockMgr as Core.LockManager
  participant InboxWB as Inbox Workbooks
  participant InvDomain as Inventory.Domain
  participant InvDB as Inventory.xlsb
  participant OutboxWB as Outbox.xlsb

  Admin->>AdminUI: Click Run Processor
  AdminUI->>Processor: RunBatch warehouseId, batchSize=500
  Processor->>LockMgr: AcquireLock INVENTORY warehouseId

  alt Lock Acquired
    LockMgr-->>Processor: TRUE expires in 3 min
    Processor->>LockMgr: UpdateHeartbeat INVENTORY
    Note over LockMgr: Set ExpiresAtUTC = Now + 3 min

    Processor->>InboxWB: Read events WHERE Status=NEW<br/>ORDER BY CreatedAtUTC LIMIT 500

    loop For each event
      Processor->>InvDomain: ApplyReceiveEvent evt

      alt Already Applied
        InvDomain->>InvDB: Check tblAppliedEvents EventID
        InvDomain-->>Processor: SKIP_DUP
        Processor->>InboxWB: UPDATE Status = SKIP_DUP
      else Apply Success
        InvDomain->>InvDB: INSERT tblInventoryLog
        InvDomain->>InvDB: INSERT tblAppliedEvents
        InvDomain-->>Processor: APPLIED
        Processor->>OutboxWB: INSERT tblOutboxEvents
        Processor->>InboxWB: UPDATE Status = PROCESSED
      else Apply Failed
        InvDomain-->>Processor: POISON ErrorCode INVALID_SKU
        Processor->>InboxWB: UPDATE Status = POISON<br/>ErrorMessage RetryCount++
      end
    end

    Processor->>LockMgr: ReleaseLock INVENTORY
    Processor->>Processor: GenerateWarehouseSnapshot
    Note over Processor: Copy snapshot to SharePoint if online
    Processor-->>AdminUI: Batch complete
  else Lock Held by Another Processor
    LockMgr-->>Processor: FALSE
    Processor-->>AdminUI: Error: Processor already running
  end
```
