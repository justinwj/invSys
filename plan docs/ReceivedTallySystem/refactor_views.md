Refactor views — Received Tally system
======================================

This file gives three focused visuals: Before/After module layout (C4 style), the runtime sequence, and a small migration plan.

View 1 — Module consolidation (Before → After)
----------------------------------------------
```mermaid
C4Component
title BEFORE: scattered receiving logic

Boundary(Tally_b, "modTS_Tally") {
  Component(TallyRecv, "Receive-side routines", "mixed")
}
Boundary(Data_b, "modTS_Data") {
  Component(DataHelpers, "Data helpers", "helpers")
}
Boundary(Launch_b, "modTS_Launchers") {
  Component(LaunchRecv, "Launch receiving", "entrypoints")
}
Boundary(Log_b, "modTS_Log") {
  Component(LogRecv, "Received logging", "logging")
}
Boundary(Form_b, "frmReceivedTally") {
  Component(Form, "Legacy form", "UI")
}
Component(ItemSearch, "frmItemSearch", "UI")

Rel(ItemSearch, TallyRecv, "calls (add/merge)")
Rel(LaunchRecv, TallyRecv, "calls")
Rel(TallyRecv, DataHelpers, "uses")
Rel(TallyRecv, LogRecv, "uses")
Rel(Form, TallyRecv, "UI hooks")
```

```mermaid
C4Component
title AFTER: consolidated in modTS_Received

Boundary(Recv_b, "modTS_Received") {
  Component(AddMerge, "AddOrMergeFromSearch()", "proc")
  Component(Aggregate, "AggregateReceived()", "proc")
  Component(ValidateWrite, "ValidateAndWrite()", "proc")
  Component(WriteInvSys, "WriteInvSys()", "proc")
  Component(WriteLog, "WriteReceivedLog()", "proc")
}
Boundary(Undo_b, "modUndoRedo (helper)") {
  Component(MUndo, "MacroUndo()", "proc")
  Component(MRedo, "MacroRedo()", "proc")
}
Component(ItemSearch2, "frmItemSearch", "UI")
Component(ConfirmBtn, "Sheet Confirm button", "UI")

Rel(ItemSearch2, AddMerge, "add/merge into ReceivedTally")
Rel(ConfirmBtn, ValidateWrite, "trigger")
Rel(AddMerge, Aggregate, "refresh aggregated view")
Rel(Aggregate, ValidateWrite, "calls")
Rel(ValidateWrite, WriteInvSys, "calls")
Rel(ValidateWrite, WriteLog, "calls")
Rel(MUndo, Aggregate, "undo staging/posted/log")
Rel(MRedo, Aggregate, "redo staging/posted/log")
```

View 2 — Runtime sequence (entry → confirm)
-------------------------------------------
```mermaid
sequenceDiagram
    participant Search as frmItemSearch
    participant RT as ReceivedTally (list)
    participant AGG as Aggregated list
    participant CNF as Confirm macro
    participant INV as invSys.RECEIVED
    participant LOG as ReceivedLog
    participant UNDO as MacroUndo
    participant REDO as MacroRedo

    Search->>RT: add/merge item (sum duplicates)
    RT->>AGG: rebuild aggregated view

    AGG->>CNF: Confirm (if happy)
    CNF-->>INV: write rows
    CNF-->>LOG: write log entry
    CNF-->>CNF: validation fails → show error (no write)

    UNDO-->>RT: undo inserts/merges
    UNDO-->>AGG: undo staging
    UNDO-->>INV: undo posted rows
    UNDO-->>LOG: undo log entry

    REDO-->>RT: redo inserts/merges
    REDO-->>AGG: redo staging
    REDO-->>INV: redo posted rows
    REDO-->>LOG: redo log entry
```

View 3 — Migration plan (flowchart)
-----------------------------------
```mermaid
flowchart TD
    classDef step fill:#dde7ff,stroke:#2f4e9c,color:#000,stroke-width:1.1px;
    classDef done fill:#dff7df,stroke:#2f6f2f,color:#000;
    classDef warn fill:#fff8d7,stroke:#b5a542,color:#000;

    S1["Move receiving routines from modTS_Tally → modTS_Received"]:::step
    S2["Move receiving entrypoints from modTS_Launchers → modTS_Received"]:::step
    S3["Inline needed helpers from modTS_Data → modTS_Received"]:::step
    S4["Move Received logging from modTS_Log → modTS_Received"]:::step
    S5["Replace frmReceivedTally with Aggregated list + Confirm button"]:::step
    S6["Update call sites to modTS_Received"]:::step
    S7["Deprecate old modules (Tally/Data/Launchers/Log)"]:::warn
    S8["Remove old modules and form"]:::done

    S1 --> S2 --> S3 --> S4 --> S5 --> S6 --> S7 --> S8
```
