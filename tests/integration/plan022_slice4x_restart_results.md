# Plan 022 Slice 4x Production clean-restart results

- Date: 2026-08-23
- Runtime: isolated generated warehouse; not dedicated-NAS acceptance
- Public entry: `mProduction.BtnOpenProductionForm`
- Final report: `reports/runtime/plan022-slice4x-restart-green2/production-reusable-production.md` (ignored machine/runtime evidence)

## D13 trace

- Focused source RED: 0/4 because no bounded restart action probe or
  two-Excel-process reusable Production path existed.
- First runtime attempt: 1/2. Runtime behavior was correct, but the new harness
  counted the two headless Domain workbooks as duplicate operator workbooks.
- Harness correction: count only station-local Production operator workbooks;
  no runtime behavior changed.
- Final GREEN: focused source contract 4/4 and packaged runtime 2/2.

## Accepted behavior

The first Excel process creates and releases the reusable multi-output fixture,
completes two batches, and saves its station-local Production workbook. The
harness then terminates Excel and starts a new process with the same isolated
warehouse and operator-workbook root. Through the public packaged launcher and
the actual Run List **Load** handler, the second process proves:

- the exact released Recipe ID/version remains available from the headless
  Designs Domain;
- the loaded reusable run session reports `Loaded=True`;
- the bound Production workbook has the same full path as in session one;
- exactly one station-local Production operator workbook exists; and
- invoking the restart probe creates no additional workbook.

The final rebuilt package remains GREEN at 74/74 packaged XLAM/restart checks
and 142/142 packaged Ribbon/compile checks. Dedicated-NAS visible Production
acceptance remains open.

The subsequent dedicated-NAS launcher run exposed a 14/16 write-on-launch RED
in the canonical Designs workbook. The idempotent schema/save correction and
final 16/16 GREEN are recorded in
`plan022_slice4x_nas_launcher_results.md`.
