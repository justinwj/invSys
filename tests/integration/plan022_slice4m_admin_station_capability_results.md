# Plan 022 Slice 4m Admin Station Capability Evidence

- Date: 2026-08-17
- Active slice: 4m — current-computer Admin capability transition
- Role/package: Core Auth consumed by Admin invSys Sign In
- Contract: after exact credential validation, copy only the same user's
  effective active legacy `S1` capabilities to the exact current-computer
  station; preserve dates, existing rows, and denies; invent no role.

## D13 RED

- Focused range 6-7: 1 passed, 1 failed.
- `TestAdminSignIn_CurrentComputerCopiesLegacyS1Capabilities`: **FAIL** — a
  valid legacy-`S1` Admin could not sign in at the computer station and no
  exact current-station capability rows appeared.
- `TestAdminSignIn_CurrentComputerDoesNotInventMissingLegacyCapability`:
  **PASS** — the pre-change code did not manufacture `ADMIN_MAINT`.

## D13 GREEN

- Focused range 6-8: 3/3, including preservation of an explicit
  current-station `DENY`.
- Core NAS/target/auth/session range 1-33: 33/33.
- Packaged XLAM: 74/74.
- Packaged Ribbon: 140/140.
- Ordered Release 1 full chain: 30/30 after a clean-session rerun. The first
  attempt completed all business and restart checks but its final read-only
  extractor encountered an orphaned headless Excel process; that process was
  terminated and the complete validator then passed.
- Plan 022 launcher contracts: 24/24; packaged no-eligible launcher state:
  3/3; Slice 4j readiness: 18/18; Slice 4l session controls: 9/9.
- Dedicated `WHT7025AE` NAS runtime: 16/16 using the separate automation user;
  the human UAT credential was neither read nor changed.
- Static baseline: expected drift RED 13/19 after source/tests changed;
  regenerated deterministic GREEN 19/19.

## Static metrics

- Components: 150
- Procedures: 4,633
- Literal `Application.Run` targets: 8
- Unresolved dynamic calls: 47
- Duplicate-body candidates: 184

## Visible acceptance remaining

Restart Excel, select `WHT7025AE`, use Admin **invSys Sign In** with the
existing warehouse user, and confirm `ADMIN_MAINT` succeeds at station
`X1-PRO-AI` and Admin controls enable.
