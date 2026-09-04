# Plan 022 Slice 4ad Process picker record-type RED

Date: 2026-08-25

Visible RED: Production Item Search opened when the operator entered an INPUT
row's **Acceptable Managed Item** cell but did not open from the corresponding
cell when Record Type was OUTPUT.

The corrected focused source contract was `5/6`: the worksheet still restricted
numbered acceptable-item targets to INPUT/REQUIREMENT. The corrected packaged
public-handler proof was `0/2`; its primary action recorded
`OutputPickerOpened=False`, `OutputPickerCommitted=False`, and
`OutputSkuRoundTrip=False`, while `OutputNameRetained=True`,
`NoPhysicalKey=True`, the INPUT picker, reusable lifecycle actions, and reusable
run remained GREEN.

This RED entered the OUTPUT **Acceptable Managed Item 1** cell through the real
worksheet selection event. It was not a compile, inventory, workbook, fixture,
or harness failure.
