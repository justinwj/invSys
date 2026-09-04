# Plan 022 Slice 4k Target Binding GREEN

- Date: 2026-08-17
- Focused current-computer and same-action ribbon tests: 2/2 passed.
- Core target/auth/write regression: 28/28 passed.
- Packaged XLAM validation: 74/74 passed.
- Packaged Ribbon validation: 142/142 passed.
- Dedicated NAS runtime validation: 12/12 passed across two clean Excel
  sessions with warehouse `WHT7025AE` and station `X1-PRO-AI`.
- Isolated Release 1 Receiving → Production → Boxing → Shipping chain: 30/30
  passed through restart and reconciliation.
- Runtime-state redaction and non-mutation suite: 62/62 passed.
- Deterministic maintenance baseline: 19/19 passed; 150 components, 4,619
  procedures, 9 literal `Application.Run` calls, 48 unresolved dynamic calls,
  and 184 duplicate-body candidates.

The selected target is now committed atomically after config validation. The
current Windows computer may enroll itself as an ordinary station when absent;
arbitrary station names remain rejected. Changing warehouse, station, or root
also signs out the prior invSys session before the new warehouse can be used.
