# Plan 022 Slice 4ah Admin inventory edit selection RED

Date: 2026-08-26

The focused contract was added before implementation and reported `0/5`.
Choosing Filtered Water from the visible **Inventory item** combo leaves its
display text in the control, but the combo Change handler rebuilds the list,
clears `ListIndex`, and never binds the selected SKU. Save therefore reports
**Choose an inventory item to edit** even though the operator selected one.

This is meaningful behavioral RED through the available form handler, not a
compile, fixture, workbook, or harness failure.
