## Phase 2 Processor Event Scope Analysis

### 1. **Shipping: Which user actions are phase-2 processor events?**

Based on the spec's **Phase 2 scope correction (3/8/26)**, phase 2 includes `RECEIVE`, `SHIP`, and `PROD` event types. However, the current Shipping flow has four distinct inventory-affecting transitions:

- `BTN_BOXES_MADE` - assembles boxes from components
- `BTN_TO_TOTALINV` - moves packaged boxes to available inventory  
- `BTN_TO_SHIPMENTS` - stages items for shipment
- `BTN_SHIPMENTS_SENT` - finalizes shipment and deducts inventory

**Recommendation:** For phase 2, you should support **only the final `BTN_SHIPMENTS_SENT` action** as the `SHIP` event. The intermediate staging actions (`BOXES_MADE`, `TO_TOTALINV`, `TO_SHIPMENTS`) are pre-posting workflow steps that prepare the shipment but don't represent the final inventory-deducting transaction. This aligns with the receive-path pattern where staging happens in UI-side tables and only the final "post" action generates an inbox event.

***

### 2. **Production: Which actions become events?**

The Production flow has three inventory-affecting actions:

- `BTN_TO_USED` - stages component usage from palette tables to invSys.USED
- `BTN_TO_MADE` - records finished goods produced (moves USED → MADE, applies WIP)
- `BTN_TO_TOTALINV` - moves finished goods from MADE → TOTAL INV (available inventory)

**Recommendation:** For phase 2, support **both `BTN_TO_MADE` and `BTN_TO_TOTALINV`** as distinct `PROD` event types:

- **`PROD_CONSUME`** for `BTN_TO_MADE` - records component usage and finished goods staging
- **`PROD_COMPLETE`** for `BTN_TO_TOTALINV` - moves finished goods to available inventory

`BTN_TO_USED` is a pre-staging action (similar to Shipping's intermediate steps) that updates the `USED` column but doesn't finalize any inventory transaction, so it should remain UI-side for phase 2.

***

### 3. **Multi-line action event granularity?**

The spec's **tblInventoryLog** schema has a 1-to-1 relationship between `EventID` (PK) and inventory delta rows. However, real-world Shipping/Production actions often generate multiple line-item deltas from one user button click.

**Recommendation:** Use **one parent event with multiple child `tblInventoryLog` rows sharing the same `EventID`**:

```vba
' Single inbox event row:
EventID = "abc-123"
EventType = "SHIP"
Payload = JSON array of line items

' Results in multiple inventory log rows:
EventID = "abc-123", SKU = "WIDGET-A", QtyDelta = -10
EventID = "abc-123", SKU = "WIDGET-B", QtyDelta = -5  
EventID = "abc-123", SKU = "WIDGET-C", QtyDelta = -3
```

This approach:
- ✅ Preserves atomic transaction semantics (all-or-nothing apply)
- ✅ Matches the existing receive-path pattern where one PO can have multiple line items
- ✅ Simplifies idempotency checks (one `EventID` lookup in `tblAppliedEvents`)
- ✅ Enables transaction-level undo (single `UndoOfEventId` link)

**Alternative (not recommended):** Creating separate child `EventID`s per line would complicate idempotency, poison handling, and undo logic, and would not align with the spec's `EventID` as primary key for both inbox and `tblAppliedEvents`.

***

### Summary

| Module | Phase 2 Processor Events | Granularity |
|--------|-------------------------|-------------|
| **Shipping** | `BTN_SHIPMENTS_SENT` only (final post) | One parent EventID → multiple inventory log rows |
| **Production** | `BTN_TO_MADE` (consume/stage) + `BTN_TO_TOTALINV` (complete) | One parent EventID → multiple inventory log rows |
| **Multi-line** | Single inbox event with JSON payload array | Multiple `tblInventoryLog` rows share one `EventID` |

This keeps phase 2 focused on the **final posting actions** that actually commit inventory changes, while deferring complex multi-stage workflows to phase 3's role UI refinement.