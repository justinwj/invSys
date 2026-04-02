Slice 1 — Distribution contract

Slice 2 — Add-ins publish verification

Slice 3 — Tester bundle writer

Slice 4 — First-run bootstrap

Slice 5 — Read-model readiness check

Slice 6 — Confirm Writes proving scenario

Slice 7 — Tester guide

Constraint block for every prompt
INVARIANTS — do not violate:
- SetupTesterStation is idempotent — safe to run multiple times on the same machine
- Auth provisioning never stores plain-text passwords or PINs
- TEST-SKU-001 seed is an admin import event through modProcessor, not a direct table write
- WH1.Receiving.Operator.xlsm is created from modRoleWorkbookSurfaces, not hand-built
- SharePoint unavailability must never block local setup completion
- CheckReceivingReadiness shows actionable messages — never raw error codes or MsgBox
- Distribution contract (docs/tester-distribution-contract.md) is the acceptance target for all slices
