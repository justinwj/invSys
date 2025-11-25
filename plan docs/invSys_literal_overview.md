invSys overview
================

- Purpose: Excel/VBA workbook that tracks inventory movements (receiving, production, shipments, adjustments) in the `invSys` table on the **INVENTORY MANAGEMENT** sheet.
- Core data: `ITEM_CODE`, `ITEM`, `UOM`, transaction columns (`RECEIVED`, `USED`, `SHIPMENTS`, `ADJUSTMENTS`, `MADE`), running totals (`TOTAL INV`, `LAST EDITED`, `TOTAL INV LAST EDIT`), and status strings (`ACTIVE`, `DEPRECATED`, `OBSOLETE`, `REMOVED`, `INACTIVE`) from `modGlobals`.
- Actions: Buttons/macros (`AddGoodsReceived_Click`, `DeductUsed_Click`, `DeductShipments_Click`, `Adjustments_Click`, `AddMadeItems_Click`) apply pending quantities, zero out entry cells, and surface user-facing messages via `DisplayMessage`.
- Audit + safety: `modUR_Transaction` wraps multi-row edits in transactions, captures snapshots (`modUR_Snapshot`), tracks undo/redo (`clsUndoAction`, `modUR_UndoRedo`), and logs changes to `InventoryLog` via `modInvLogs` with GUID-based `LOG_ID`s.
- UI: UserForms for item search (`frmItemSearch` uses `gSelectedCell`), admin/user management, production/receiving/shipping tally, and recipe/substitution helpers.
- Tooling: `modExportImportAll` syncs VBA components to disk and back for source control; `modDiagnostics` and `modErrorHandler` provide lightweight tracing and error messaging hooks.
- Integration points worth knowing: `modGlobals.GetItemUOMByRowNum` resolves default UOM from row/item/name; `modUR_ExcelIntegration` and `modUR_Snapshot` manage workbook-level interactions (snapshots, IDs); `modDynamicForms`/`modDiagramCore` support dynamic UI wiring.
