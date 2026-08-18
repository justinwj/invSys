# Demo Inventory Lifecycle Projection RED

- Date: 2026-08-17
- Status: **FAIL**
- Callback: `modAdmin.Seed_DemoInventory`
- Runtime: isolated generated test warehouse
- Existing lifecycle behavior: form actions, idempotent repeated Seed, exact-key Delete,
  validated Upload, repeated Upload, and destructive guards all passed.
- Failing behavior: after Delete, Inventory Viewer and Receiving still rendered depleted
  zero-quantity demo groups; after Upload, those stale zero groups remained alongside the
  one active uploaded group.

This is the focused behavioral RED protecting the rule that current-inventory operator
projections suppress groups whose summed available quantity is zero or less while the
canonical entity/audit history remains intact.

## Dataset-selection RED

After the user clarified that the seed dataset must be selectable, the same packaged
callback was rerun with a validated CSV path supplied to the **Seed Demo Inventory**
action. It failed because Seed ignored the selected path and applied the built-in kit.
The source form check was 8/9: no dataset selector, built-in-kit choice, or retained
uploaded-path property existed yet.
