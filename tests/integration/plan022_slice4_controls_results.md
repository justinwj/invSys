# Plan 022 Slice 4 final controls results

**Date:** 2026-08-10  
**Slices:** 4d, 4h, and 4i

## Outcome

The packaged Release 1 forms now use a visible Inventory Viewer ribbon icon,
resize all Box Builder and Box Maker lists, keep list headers aligned, filter
Box Builder component choices locally, display `NA` for non-versioned managed
items, show Receiving Entries History separately from the managed-item selector,
and preserve exact string `System_Key` identity through the reachable Shipping
form, reservation, staging, and sent-event path.

## Focused RED and GREEN

The new `Test-R1FinalControlAcceptance.ps1` contract initially recorded 0/12:
the Viewer icon, Boxing anchors/search/headers/`NA`, Receiving history split,
and Shipping `System_Key` form/event path were absent. After implementation it
is 12/12 GREEN. The packaged Shipping public-launcher proof additionally enters
the real form and verifies resize geometry, local search (one of two synthetic
rows remains), `NA` display, aligned headers, string-key preservation, and a
reservation key derived from that string. The packaged Receiving public
launcher loads two synthetic `ReceivedLog` entries through the form Refresh
handler while preserving captured-workbook binding.

## Regression evidence

| Evidence | Result |
|---|---:|
| R1 final control contract | 12/12 |
| Slice 10 Receiving stabilization | 10/10 |
| Slice 11 Shipping/Boxing stabilization | 11/11 |
| Shipping status anchor | 4/4 |
| Packaged Shipping layout/search/identity | 1/1 |
| Packaged Receiving history durability | 1/1 |
| Packaged XLAM suite | 74/74 |
| Packaged RibbonX suite | 136/136 |

The full five-package build completed successfully with Excel closed before the
build. Runtime reports under `reports/runtime/` remain ignored machine evidence;
the deterministic source/result files above are the committed evidence.

## Static maintenance evidence

The deterministic baseline timestamp is `2026-08-10T20:20:00Z`. Components
remain 149. Procedures increased from 4,566 to 4,585 for the new form/history
loaders and packaged seams. Literal `Application.Run` targets remain 9,
unresolved dynamic calls remain 48, duplicate-body candidates remain 185, and
same-project late-binding candidates remain 8.

Two existing oversized components have an explicit Slice 4 exception:

- `frmShipmentsTally` grew 207 lines for the two Boxing page layouts, shared
  header alignment, component search/`NA` display, `System_Key` controls, and
  public packaged layout/search/identity proof.
- `modTS_Shipments` grew 128 lines for the reachable `System_Key` form-action,
  reservation/event backing path and packaged adapters.

This is bounded Release 1 contract work, not an exception for new unrelated
features. Duplicate-body and dynamic-dispatch ratchets did not regress. Private
legacy worksheet-maintenance routines remain review candidates and are not
deletion-authorized by the scanner.
