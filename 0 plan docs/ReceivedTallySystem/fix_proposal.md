Tally Received fix proposal (Mermaid, updated)
==============================================

Goal: remove the legacy “Tally Received” form/button; use the Aggregated list (auto-built) and a generated Confirm button instead.

```mermaid
flowchart LR
    classDef table fill:#e8f9ff,stroke:#2c7a9b,color:#000,stroke-width:1.2px;
    classDef proc fill:#dde7ff,stroke:#2f4e9c,color:#000,stroke-width:1.2px;
    classDef note fill:#fff8d7,stroke:#b5a542,color:#000,stroke-dasharray:3 3;
    classDef data fill:#dff7df,stroke:#2f6f2f,color:#000;
    classDef log fill:#f7eadb,stroke:#8c6239,color:#000,stroke-dasharray:4 3;

    U["User selects item in frmItemSearch"]:::proc
    RT["ReceivedTally\n(REF_NUMBER, ITEMS, QUANTITY, PRICE)"]:::table
    AGG["Aggregated list\n(REF_NUMBER, ROW, ITEM_CODE, VENDOR(s), VENDER_CODE, DESCRIPTION, ITEM, UOM, LOCATION, PRICE, QUANTITY)"]:::table
    CNF["Generated button: Confirm writes"]:::proc
    INV["invSys (inventory)"]:::data
    RLOG["ReceivedLog"]:::log

    U -->|add/merge| RT -->|"auto-aggregate (sum qty + refs; split if fields differ)"| AGG -->|Confirm| CNF
    CNF -->|write rows| INV
    CNF -->|write log| RLOG

    %% Fix notes
    F1["Remove legacy Tally Received button/form; use Aggregated list as the staging view with generated Confirm button."]:::note
    F2["Shared helpers: ColumnIndex/FindRowByValue must be Public/in shared module to avoid compile errors in aggregation routines."]:::note
    F3["UOM lookup: use invSysData_Receiving.TALLY_ROW (row num) when calling GetUOMFromInvSys; keep ITEM/ITEM_CODE fallback."]:::note
    F4["Confirm ordering: write invSys first, then ReceivedLog; capture undo, then clear staging tables."]:::note

    RT --- F1
    AGG --- F2
    AGG --- F3
    CNF --- F4
```
