# Plan 022 Slice 4s Shipping Exact-Key Results

Date: 2026-08-20

## Contract and RED

The visible `WHT7025AE` checkpoint selected a shippable box with an immutable
`System_Key`, then Shipping **Add** failed with `invSys table missing TOTAL
INV/SHIPMENTS/ROW columns.` Box Designer also exposed zero-balance durable
entities as selectable repeated component rows. Required Shipping/Boxing saves
surfaced repeated native Saving notices.

The focused test was created and run before implementation. Initial result:

- public Shipping Add path: PASS;
- current-schema exact-key reserve apply: FAIL;
- active exact-entity component projection: FAIL;
- distinct positive identity preservation: FAIL;
- Shipping action quiet boundaries: FAIL; and
- status-bar restoration: FAIL.

## GREEN

- `tests/tooling/Test-Plan022Slice4sShippingExactKey.ps1`: 6/6 PASS.
- Phase 6 Excel tests 134-136: 3/3 PASS, including the same public
  `ShipmentsFormCommitLine` action on a `ROW`-free workbook.
- Shipping/Boxing stabilization: 11/11 PASS.
- R1 final control acceptance: 12/12 PASS.
- Plan 022 workflow readiness: 18/18 PASS.
- Deployed live role workflows: 46/46 PASS.
- Ordered Release 1 full chain: 30/30 PASS.
- Deterministic maintenance baseline: 19/19 PASS.

The current-schema test proves Shipping Add preserves the selected string
`System_Key` and applies the exact entity's local `SHIPMENTS` reservation.
Component choices now omit nonpositive balances and deduplicate only repeated
projections of the same key. Required saves remain authoritative; the quiet UI
boundary now also hides and restores Excel's status bar.

Static evidence: 150 components, 4,702 procedures, 965 scanner candidates, 8
literal `Application.Run` targets, 47 unresolved dynamic expressions, and 184
duplicate-body groups.

## Non-gating legacy inspector finding

The generic visible package inspector opened all five add-ins and passed
32/34. Its two failures expected retired `AggregateBoxBOM_Log` and
`AggregatePackages_Log` support worksheets. The active deployed form/action
gate is the 46/46 live-role suite.

## Visible acceptance still required

Against `WHT7025AE`, confirm Box Designer omits zero-balance and same-key
duplicate component rows, Shipping Add no longer reports a `ROW` requirement,
and record whether native Saving notifications remain visible during Add, Save
Box, Make/Unmake, and Shipments Sent.
