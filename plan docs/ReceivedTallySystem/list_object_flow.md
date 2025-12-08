List-object-based received tally flow (no runtime form) — Mermaid views
======================================================================

This is split into two focused views: (A) staging with merge-on-insert and confirm trigger; (B) confirm/undo/redo with logging/error. The separate lookup step was removed—aggregation handles lookup/rollup.

View A — Received Tally system (staging + confirm trigger)
---------------------------------------------------------
```mermaid
flowchart LR
    classDef button fill:#dde7ff,stroke:#2f4e9c,color:#000,stroke-width:1.3px;
    classDef list fill:#e8f9ff,stroke:#2c7a9b,color:#000,stroke-width:1.3px;
    classDef note fill:#fff8d7,stroke:#b5a542,color:#000,stroke-dasharray:3 3;

    SRC["Item Search<br/>Form"]:::button
    RT["ReceivedTally<br/>list object"]:::list
    BTN["Confirm writes"]:::button
    AGG["Aggregated tally list<br/>(new list object)"]:::list

    NOTE1["On add from Item Search: merge identical items (same ITEM/UOM or ITEM_CODE) by summing QTY into an existing row; otherwise append a new row."]:::note
    NOTE2["Aggregated list is the staging view used for confirm; columns: ITEM, QTY, UOM, PRICE, ITEM_CODE, ROW."]:::note

    SRC -->|select item<br/>add/merge| RT
    RT -->|auto-aggregate| AGG
    AGG -->|"Confirm (if happy)"| BTN

    RT --- NOTE1
    AGG --- NOTE2
```

View B — Confirm / Undo / Redo / Logging / Error
------------------------------------------------
```mermaid
flowchart LR
    classDef button fill:#dde7ff,stroke:#2f4e9c,color:#000,stroke-width:1.3px;
    classDef list fill:#e8f9ff,stroke:#2c7a9b,color:#000,stroke-width:1.3px;
    classDef data fill:#dff7df,stroke:#2f6f2f,color:#000;
    classDef log fill:#f7eadb,stroke:#8c6239,color:#000,stroke-dasharray:4 3;
    classDef note fill:#fff8d7,stroke:#b5a542,color:#000,stroke-dasharray:3 3;
    classDef error fill:#ffe6e6,stroke:#b34747,color:#000,stroke-width:1.2px,stroke-dasharray:3 2;

    AGG["Aggregated tally list<br/>(new list object)"]:::list
    CNF["Confirm writes"]:::button
    INV["invSys.RECEIVED"]:::data
    RLOG["ReceivedLog"]:::log
    UNDO["Undo/edit<br/>then re-tally"]:::button
    REDO["Redo last<br/>macro"]:::button
    ERR["Error/validation<br/>failure shown<br/>(no write)"]:::error

    NOTE3["Confirm: validate; write invSys.RECEIVED + ReceivedLog; clear staging tables."]:::note
    NOTE4["Undo/edit: adjust tally rows, re-tally."]:::note
    NOTE5["Redo: reapply last undone macro (staged rows and posted rows/log)."]:::note
    NOTE6["Logging and error handling: successes log; failures surface to the user and log an error, no data write."]:::note

    AGG -->|"Confirm (if happy)"| CNF
    CNF -->|write rows| INV
    CNF -->|write log entry| RLOG
    CNF -->|validation fails| ERR
    ERR -.->|log error| RLOG

    UNDO -.->|undo aggregated list| AGG
    UNDO -.->|undo posted rows| INV
    UNDO -.->|undo log entry| RLOG

    REDO -.->|redo staging| AGG
    REDO -.->|redo posted rows| INV
    REDO -.->|redo log entry| RLOG

    AGG --- NOTE3
    UNDO --- NOTE4
    REDO --- NOTE5
    RLOG --- NOTE6
```
