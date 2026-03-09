Event Entry Points — Mermaid (sequence view)
============================================

Buttons (Confirm / Undo / Redo) are generated in `modTS_Received`. This view shows who triggers what, without layout clutter.

```mermaid
sequenceDiagram
    participant Search as frmItemSearch (Add_Click)
    participant RT as ReceivedTally (list)
    participant AGG as Aggregated list
    participant CNF as Confirm button macro
    participant INV as invSys.RECEIVED
    participant LOG as ReceivedLog
    participant UNDO as Undo button macro
    participant REDO as Redo button macro

    Search->>RT: add/merge item (sum qty, concat refs)
    RT->>AGG: rebuild aggregated view

    AGG->>CNF: Confirm (if happy)
    alt validation OK
        CNF-->>INV: write rows
        CNF-->>LOG: write log entry
    else validation fails
        CNF-->>CNF: show error (no write)
    end

    UNDO-->>RT: undo inserts/merges
    UNDO-->>AGG: undo staging
    UNDO-->>INV: undo posted rows
    UNDO-->>LOG: undo log entry

    REDO-->>RT: redo inserts/merges
    REDO-->>AGG: redo staging
    REDO-->>INV: redo posted rows
    REDO-->>LOG: redo log entry
```
