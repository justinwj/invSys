# Plan 022 Deployed Operations Launcher and NAS Runtime Evidence

- Date: 2026-07-28
- Package set: R1-5
- Dedicated NAS leaf: `Plan022-R1-UAT-20260728-7025AE7B`
- Root fingerprint: `51101D8F55ED`
- Warehouse/station: `WHT7025AE` / `S1`
- Automated status: GREEN except known Shipping D14 `ROW` blocker
- User acceptance status: RED

## 2026-08-04 operator checkpoint

The user returned the visible dedicated-NAS checkpoint against the prepared
test warehouse/station.

Accepted observations:

- repeated Receiving, Production, and Shipping launches reused their
  respective operator workbooks; no additional role workbook opened;
- forms remained bound when another workbook was activated and minimized with
  their owning workbook; and
- launcher behavior remained stable after Excel restart.

Failed observations:

- the visible Admin Seed Demo Inventory action returned a report equivalent to
  `Demo inventory seeded.|Applied=1|Processor=Applied=1; SkipDup=0; Poison=0`,
  but none of the three demo entities appeared after refresh; the session RunId
  is intentionally omitted from committed evidence; and
- the Production form's native window maximized while the MultiPage and child
  controls remained at a small base-size footprint in the upper-left corner.

A read-only inspection of the saved station-local operator workbook packages
after the checkpoint found 18 nonblank `System_Key` rows in each role's
`invSys` table. Sixteen rows used supported demo SKU codes and all 16 recorded
`Condition=GOOD`; the three current seed codes were present five times each,
consistent with repeated seed actions creating new durable keys. No workbook
was opened, refreshed, saved, or closed for this inspection. This narrows the
remaining Seed failure to the operator-visible form refresh/list contract (or
the acceptance presentation of those rows), rather than absence of canonical
or saved operator projection data. The user's visible failure remains RED.

Source review after the checkpoint also found active Shipping `ROW` labels and
fields in `frmShipmentsTally`, contrary to D14. The checkpoint therefore
advances launcher reuse/binding/restart evidence but does not complete Plan 022
or Release 1 acceptance.

## D13 trace

The meaningful packaged callback RED reproduced the three reported launcher
failures before implementation:

- Receiving returned `Open a Receiving operator workbook before using the
  Receiving form.`;
- Production failed with `Type mismatch`; and
- Shipping failed with `Type mismatch`.

The preserved redacted RED report is
`reports/runtime/plan022-slice0/packaged-launcher-red-before-diagnostics.md`
(SHA-256
`7A8DCE899219F443A2F7BDCF60B818E5DE5798BA89AF3F5FCE2B010516884063`).

The first dedicated NAS run then exposed two additional behavioral REDs:
Shipping created two modeless forms on repeated launch, and launcher
authorization rewrote canonical Config/Auth workbooks. The first corrective
pass made those boundaries GREEN.

The first user checkpoint then supplied a second meaningful RED batch:

- Production returned the missing-workbook instruction instead of creating its
  station-local workbook;
- Shipping returned the same missing-workbook instruction; and
- closing the Receiving form left a disappeared modeless-form proxy whose next
  launch returned an automation error.

The expanded focused suite recorded 17 PASS / 4 RED before the second
correction. A further packaged RED proved that a newly generated Shipping
workbook needed a visible landing sheet. The final focused suite is 24/24.
Receiving now invalidates its cached launcher reference in `UserForm_QueryClose`;
all three roles share the Core-owned open/create primitive; and automated NAS
validation uses a distinct account so it cannot overwrite the prepared human
UAT PIN.

The first downstream Admin control checkpoint on 2026-08-02 exposed another
meaningful packaged RED: `modAdmin.Seed_DemoInventory` remained in its broad
warehouse-directory/context path for more than 45 seconds before the selection
form could complete, matching the operator-observed flashing followed by an
application/object-defined error. The callback was resolving the active
canonical Config workbook as an Admin surface and scanning remembered roots
even though a valid current warehouse target was already selected.

The 2026-08-03 correction makes the Seed Demo Inventory callback request only
the current selected target. General View Warehouses scanning remains
unchanged. The callback now reports its failing stage, error number, sanitized
source, and description, and the packaged automation seam injects only the
form selection before calling the same public callback. Preserved evidence:

- `tests/integration/admin_seed_callback_red_results.md`;
- `tests/integration/admin_seed_callback_green_results.md`; and
- `tools/validate_admin_seed_inventory_callback.ps1`.

After the 2026-08-04 visible checkpoint, the focused callback validator was
expanded to require the same three unique `System_Key` values and
`Condition=GOOD` across canonical inventory, the published snapshot, a saved
Receiving operator workbook, and the Receiving form's actual Refresh click
handler/list rendering. The final deployed package passed all of those checks,
including three visible `DEMO-` rows through the real Refresh handler. This
does not erase the operator's failed visible checkpoint; it proves the seeded
entities are available when the specified Receiving inventory control is used.

The Production layout validator was expanded to require `Zoom=100` and at
least 90% DPI-adjusted native client fill after maximize. Its focused RED
captured the packaged form at `Zoom=60`, reproducing the shrunken-control
symptom. Removing the DPI-derived form zoom produced GREEN across all four
pages, three supported sizes, and minimize/restore/maximize/restore. The
maximized form measured approximately 1451 x 875 points against a 1440 x 847
point client area, with zero out-of-bounds controls and zero interactive
overlaps.

## Automated GREEN

| Evidence | Result |
|---|---:|
| Plan 022 focused launcher contracts | 24/24 |
| Final packaged launcher matrix | 17/17 |
| Dedicated NAS runtime, two clean Excel sessions | 12/12 |
| Tool B contracts, redaction, and no-mutation proof | 62/62 |
| Source Ribbon generation contract | 46/46 |
| Packaged five-XLAM compile/surface/restart | 54/54 |
| Packaged RibbonX | 136/136 |
| Live role workflow | 46/46 |
| Ordered Release 1 full chain | 30/30 |
| Packaged Admin Seed Demo Inventory callback and Receiving Refresh action | GREEN; 3 unique D14 entities with `Condition=GOOD` match across canonical inventory, snapshot, saved operator workbook, and 3 visible filtered rows |
| Production native resize | GREEN; 3 sizes x 4 pages, `Zoom=100`, native minimize/restore/maximize/restore, client fill PASS, no out-of-bounds controls or interactive overlaps |
| Create Warehouse / repeated seed D14 lifecycle | 15/15 |
| Slice 5 behavior locks | 13/13 |
| Slices 6, 8, 9, 10, 11, 12, 13, and 14 static locks | 10/10, 14/14, 8/8, 10/10, 11/11, 11/11, 14/14, 9/9 |

The final dedicated NAS evidence is
`reports/runtime/plan022-nas/dedicated-nas-runtime.md`. In each clean session:

- the exact five approved manifest hashes were loaded;
- the selected dedicated NAS target resolved to `WHT7025AE` / `S1`;
- Receiving, Production, and Shipping callbacks each completed twice;
- each role self-provisioned exactly one workbook from an empty isolated local
  operator root;
- exactly one modeless form per role remained;
- zero canonical files changed merely from launcher use; and
- read-only extraction invoked zero mutations and changed no inspected hash.

The account startup registry contains only the consolidated Operations and
Admin leaf add-ins. The final manifest contains five packages with zero hash
mismatches.

## Static maintenance review

The regenerated baseline records:

- components: 164;
- procedures: 4,562;
- maintenance candidates: 964;
- reviewed candidates: 966;
- duplicate-body groups: 187 -> 187;
- literal `Application.Run` targets: 9; and
- unresolved dynamic calls: 48.

Relative to the immediately preceding committed baseline, the two added
procedures are the Receiving form-action test seam and its packaged wrapper.
They are required to exercise the same public Refresh handler used by the
operator. The scanner consequently adds one review-only reachability candidate;
it is protecting test infrastructure and is not deletion-authorized. Component,
duplicate-body, `Application.Run`, and unresolved-dynamic-call counts did not
increase.

## Remaining completion gate

Plan 022 is not complete until the operator verifies existing demo rows through
the Receiving inventory controls and repeats Production maximize/restore on the
new package. Shipping's separate D14 `ROW` identity conflict also remains open
and requires its own focused correction before Release 1 acceptance.
