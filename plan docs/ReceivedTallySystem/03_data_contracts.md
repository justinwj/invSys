Data Contracts (tables/lists) — Mermaid
=======================================

Lookup, staging, and posting relationships for the Received Tally system. Prices/taxes are not tracked; logging captures qty/refs.

Flowchart
---------
```mermaid
flowchart TB
    classDef list fill:#e8f9ff,stroke:#2c7a9b,color:#000,stroke-width:1.2px;
    classDef data fill:#dff7df,stroke:#2f6f2f,color:#000;
    classDef log fill:#f7eadb,stroke:#8c6239,color:#000,stroke-dasharray:4 3;
    classDef note fill:#fff8d7,stroke:#b5a542,color:#000,stroke-dasharray:3 3;

    INV_SRC["invSys (catalog source)\nROW, ITEM_CODE, VENDOR(s), LOCATION, ITEM, UOM, RECEIVED, TOTAL INV, …"]:::data
    RT["ReceivedTally\n(REF_NUMBER, ITEMS, QUANTITY)\nFast entry: ITEM text + QTY only"]:::list
    AGG["AggregateReceived\n(REF_NUMBER, ITEM_CODE, VENDORS, VENDOR_CODE, DESCRIPTION, ITEM, UOM, QUANTITY, LOCATION, ROW)\nMerge: same ITEM/UOM or ITEM_CODE → sum QUANTITY + concat REF_NUMBERs; if other fields (e.g., LOCATION/VENDOR) differ, keep separate rows"]:::list
    INV["invSys (inventory target)\n(ROW, ITEM_CODE, ITEM, UOM, RECEIVED, TOTAL INV, …, TIMESTAMP)"]:::data
    RLOG["ReceivedLog\n(REF_NUMBER, ITEMS, QUANTITY, UOM, VENDOR, LOCATION, ITEM_CODE, ROW, SNAPSHOT_ID, ENTRY_DATE)"]:::log

    INV_SRC -->|"Item Search lookup (pre-existing items only)"| RT
    INV_SRC -->|"Item Search populates detail fields (ROW, ITEM_CODE, VENDOR/LOCATION, etc.)"| AGG
    RT -->|"auto-aggregate (sum qty, combine refs; split if fields differ)"| AGG
    AGG -->|"Confirm write (add qty to RECEIVED; item must already exist)"| INV
    AGG -->|"Confirm write"| RLOG

    NOTE1["Constraint: Confirm only adds QUANTITY to existing invSys rows; items must pre-exist in invSys to appear in Item Search."]:::note
    NOTE2["Item Search writes ITEM into ReceivedTally; detail fields (ROW, ITEM_CODE, LOCATION, VENDOR, etc.) come from invSys into AggregateReceived so Confirm can target the correct invSys row."]:::note
    NOTE3["ReceivedLog records the transaction (qty, refs, location/vendor, snapshot)."]:::note

    INV --- NOTE1
    RT --- NOTE2
    RLOG --- NOTE3
```

Block diagram (distinctions: merge logic and flows)
---------------------------------------------------
```mermaid
block-beta
columns 3
  INV_SRC["invSys (catalog source)\nROW, ITEM_CODE, VENDOR(s), LOCATION, ITEM, UOM, RECEIVED, TOTAL INV, …"]
  space
  RT["ReceivedTally\nREF_NUMBER, ITEMS, QUANTITY"]
  space
  AGG["AggregateReceived\nREF_NUMBER, ITEM_CODE, VENDORS, VENDOR_CODE, DESCRIPTION, ITEM, UOM, QUANTITY, LOCATION, ROW"]
  blockArrowINVRT<["Item Search lookup\n(pre-existing items only)"]>(right)
  blockArrowINVAGG<["populate detail fields\n(ROW, ITEM_CODE, vendor/location)"]>(right)
  blockArrowRTAGG<["auto-aggregate<br/>(sum qty + concat refs;<br/>split if fields differ)"]>(right)
  blockArrowAGGINV<["write rows<br/>(add qty to RECEIVED; item must exist)"]>(right)
  blockArrowAGGRLOG<["write log<br/>(with snapshot)"]>(down)
  INV["invSys (inventory target)\nROW, ITEM_CODE, ITEM, UOM, RECEIVED, TOTAL INV, …, TIMESTAMP"]
  RLOG["ReceivedLog\nREF_NUMBER, ITEMS, QUANTITY, UOM, VENDOR, LOCATION, ITEM_CODE, ROW, SNAPSHOT_ID, ENTRY_DATE"]
  style INV_SRC fill:#dff7df,stroke:#2f6f2f,stroke-width:2px
  style RT fill:#e8f9ff,stroke:#2c7a9b,stroke-width:2px
  style AGG fill:#e8f9ff,stroke:#2c7a9b,stroke-width:2px
  style INV fill:#dff7df,stroke:#2f6f2f,stroke-width:2px
  style RLOG fill:#f7eadb,stroke:#8c6239,stroke-width:2px,stroke-dasharray:4 3
```

Entity Relationship view
------------------------
```mermaid
erDiagram
    INV_SRC ||--o{ RECEIVEDTALLY : "lookup (items must pre-exist)"
    RECEIVEDTALLY ||--o{ AGGREGATE_RECEIVED : "aggregates (sum qty + concat refs; split if fields differ)"
    AGGREGATE_RECEIVED ||--|{ INV : writes (add qty to RECEIVED)
    AGGREGATE_RECEIVED ||--|{ RLOG : logs

    INV_SRC {
        string ROW
        string ITEM_CODE
        string ITEM
        string UOM
        string LOCATION
        decimal RECEIVED
        decimal TOTAL_INV
    }
    RECEIVEDTALLY {
        string REF_NUMBER
        string ITEMS
        decimal QUANTITY
    }
    AGGREGATE_RECEIVED {
        string REF_NUMBER
        string ITEM_CODE
        string VENDORS
        string VENDOR_CODE
        string DESCRIPTION
        string ITEM
        string UOM
        decimal QUANTITY
        string LOCATION
        string ROW
        string REF_NUMBERS
    }
    INV {
        string ROW
        string ITEM_CODE
        string ITEM
        string UOM
        decimal RECEIVED
        decimal TOTAL_INV
        string LOCATION
        datetime TIMESTAMP
    }
    RLOG {
        string REF_NUMBER
        string ITEMS
        decimal QUANTITY
        string UOM
        string VENDOR
        string LOCATION
        string ITEM_CODE
        string ROW
        string SNAPSHOT_ID
        datetime ENTRY_DATE
    }
```
