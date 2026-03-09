Contract answers for your Phase 2 questions:

## 1. OccurredAtUTC vs AppliedAtUTC Timing

**Yes, treat `OccurredAtUTC` as the inbox row's `CreatedAtUTC`**. The spec shows inbox tables have `CreatedAtUTC` (when the event was written to inbox), and `tblInventoryLog` has both `OccurredAtUTC` (business event time) and `AppliedAtUTC` (processor application time). For the first processor slice, the cleanest interpretation is:

- **Inbox**: `CreatedAtUTC` = when user/role UI created the event row
- **During Apply**: `InventoryDomain.Apply` sets `OccurredAtUTC = inbox.CreatedAtUTC` and `AppliedAtUTC = Now()` when writing to `tblInventoryLog`

This maintains event-sourcing semantics where occurrence time is preserved from creation, and application time tracks actual processor execution.

## 2. Processor User Identity for Auth Checks

**The processor should run as a fixed service/admin identity**. The existing auth fixture already includes `INBOX_PROCESS` capability, indicating processor operations use a dedicated service account. The processor validates that the original event creator had proper permissions when the event was written, but the processor itself runs under an elevated service identity that has `INBOX_PROCESS` capability across all warehouses.

From the workflow diagram in the spec, `Core.Processor` interacts with `LockMgr` and domain modules with warehouse-level scope, not station-level user credentials. The processor should not require an explicit operator `UserId` from Admin UI. It uses a configured service account identity loaded from `Core.Config`.

## 3. Phase 2 Fixture Generation Strategy

**Use the same script-based approach as Phase 1**. The spec explicitly shows Phase 1 used `create_phase1_fixture_xlsx.ps1` in the tools directory, and the testing philosophy emphasizes reproducible, version-controlled test fixtures. For Phase 2, create:

- `tools/create_phase2_fixture_xlsx.ps1` to generate:
  - `invSys.Inbox.Receiving.S1.xlsb` with sample inbox events
  - `invSys.Inbox.Shipping.S1.xlsb` with sample inbox events
  - `invSys.Inbox.Production.S1.xlsb` with sample inbox events
  - `WH1.invSys.Data.Inventory.xlsb` with base schema and optional seed data

This keeps fixture generation logic explicit, allows CI/CD integration, and prevents binary workbook drift in the repository. The spec's testing strategy emphasizes automation and reproducibility.

## 4. Phase 2 Test Scope and Excel Automation

**Land VBA unit tests plus a manual harness first**. The Phase 2 deliverables explicitly state:

**Tests:**
- Test: `AcquireLock/ReleaseLock` + heartbeat lifecycle
- Test: Manual receiving, shipping, and production inbox rows -> Run processor -> verify `tblInventoryLog` and `tblAppliedEvents`
- Test: Duplicate EventID marked `SKIP_DUP`

This indicates manual/semi-automated validation initially. The progression from the spec shows:
1. **Phase 2**: Core processor mechanics with manual workbook validation
2. **Phase 3**: Role UI with full Excel automation (`UI -> Create events -> Process -> Verify domain logs`)

Wire **basic Excel automation for Phase 2** (opening workbooks, reading/writing tables, running processor), but defer full end-to-end role-UI-driven flows until Phase 3 when the UI modules exist. Phase 2 should still cover processor/domain handling for all role event types (`RECEIVE`, `SHIP`, `PROD`); what remains deferred is the UI-driven creation flow. This matches the spec's incremental validation approach where Phase 2 focuses on processor stability before adding UI complexity.

Your phase plan is solid. The spec confirms all three Phase 1 modules landed cleanly and Phase 2 is correctly scoped around `Core.LockManager`, `Core.Processor`, and `InventoryDomain.Apply` across `RECEIVE`, `SHIP`, and `PROD` event paths.
