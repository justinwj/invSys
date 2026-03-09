# invSys VBA Undo/Redo – proposal

> Scope: VBA-only, workbook-local, covers value/formula writes plus listobject row add/delete in invSys sheets. Ignores native Excel undo stack and any external/import actions.

## Goals / requirements
- One centralized undo/redo for all VBA writes in invSys.
- Works across modules; actions are pushed when macros run.
- Treat “last user action” as the unit of undo (e.g., a Confirm or a multi-cell fill).
- Redo re-applies exactly what the last undo reversed.
- In-memory only; cleared on close.

## Constraints / assumptions
- We do **not** attempt to restore Excel native undo.
- Only ranges/listobjects in this workbook.
- Structural support limited to:
  - ListObject row add/delete
  - Cell value/formula changes
  - (Optional later) formats/validation
- Large bulk ops may use a per-action snapshot (table-sized), not full workbook.

## Data model
- **ActionStack**
  - UndoStack (LIFO)
  - RedoStack (LIFO)
- **Action (clsUndoAction)**
  - ActionId, Timestamp
  - Deltas (collection)
  - Optional Snapshot (for big ops)
- **Delta**
  - Address or LO row/col (supports row add/delete)
  - BeforeValue / AfterValue
  - OpType: `CellValue`, `RowAdd`, `RowDel`, `TableResize` (optional)

## API (modUndo)
- `BeginAction(label)` – clears temp buffer.
- `LogDelta(addressOrRowInfo, beforeVal, afterVal, opType)`
- `CommitAction()` – push to Undo, clear Redo.
- `AbortAction()` – discard temp buffer.
- `Undo()` – pop Undo, apply reverse deltas, push to Redo.
- `Redo()` – pop Redo, apply forward deltas, push to Undo.
- Helpers:
  - `LogRowAdd(lo, rowIndex, rowValues)`
  - `LogRowDel(lo, rowIndex, rowValues)`
  - `ApplyDelta(delta, direction)`

## Integration points (where to call)
- ConfirmWrites (receiving): wrap the whole confirm in Begin/Commit; log invSys RECEIVED changes and ReceivedLog inserts.
- MergeIntoAggregate / quantity sync: if treated as staging, you can skip logging; if you want undo at staging level, wrap add/merge.
- Any other macro that writes to invSys tables.

## Edge cases
- Table has formulas: store/reapply formulas for cells with formulas (small map).
- Row delete while table filtered: use ListRow.Index on visible rows; restore via stored values.
- Mixed value types: store Variant as-is; preserve numberformat optionally.

## Diagrams

### Flow – action lifecycle
```mermaid
flowchart LR
  B[BeginAction] --> L[LogDelta n times]
  L -->|Commit| C[CommitAction: push Undo; clear Redo]
  L -->|Abort| A[AbortAction: discard buffer]
  U[Undo] --> ApplyR[Apply reverse deltas]
  ApplyR --> PushR[Push to Redo]
  R[Redo] --> ApplyF[Apply forward deltas]
  ApplyF --> PushU[Push to Undo]
```

### Sequence – user confirm & undo
```mermaid
sequenceDiagram
  participant User
  participant Confirm as ConfirmWrites
  participant UndoSys as modUndo
  User->>Confirm: Click Confirm
  Confirm->>UndoSys: BeginAction("Confirm")
  Confirm->>UndoSys: LogDelta(...) x N
  Confirm->>UndoSys: CommitAction()
  User->>UndoSys: Undo
  UndoSys->>UndoSys: Apply reverse deltas
  UndoSys->>UndoStack: Move to Redo
```

### Class sketch
```mermaid
classDiagram
  class clsUndoAction{
    +ActionId
    +Timestamp
    +Deltas : Collection
    +Snapshot : Variant
    +AddDelta()
    +Apply(direction)
  }
  class modUndo{
    +BeginAction()
    +LogDelta()
    +CommitAction()
    +AbortAction()
    +Undo()
    +Redo()
  }
  clsUndoAction --> modUndo : used by
```

### State (stacks)
```mermaid
stateDiagram-v2
  [*] --> Idle
  Idle --> Active : BeginAction
  Active --> Idle : AbortAction
  Active --> Idle : CommitAction(push Undo, clear Redo)
  Idle --> Idle : Undo(pop Undo, push Redo)
  Idle --> Idle : Redo(pop Redo, push Undo)
```

## Next steps
- Decide granularity: log only confirmed writes (recommended) vs. every staging edit.
- Implement clsUndoAction (deltas) and modUndo (stacks & API).
- Wire ConfirmWrites to Begin/Log/Commit; wire buttons to Undo/Redo.
- Optional: add row add/delete helpers for listobjects.
